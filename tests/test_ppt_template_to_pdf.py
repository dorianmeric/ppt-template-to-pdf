"""Tests for ppt_template_to_pdf.

Only the pure functions and the PPTX path are covered. PDF export needs a real
PowerPoint install and is exercised manually.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

import ppt_template_to_pdf as subject


# --- helpers ------------------------------------------------------------


def quiet() -> subject.Reporter:
    return subject.Reporter(quiet=True)


def build_template(path: Path) -> Path:
    """Create a small deck with one named shape and one token shape."""
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    named = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    named.name = "CLIENT_NAME"
    named.text_frame.text = "old client"

    token = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
    token.text_frame.text = "{{ SUMMARY }}"

    presentation.save(str(path))
    return path


def fill_and_reopen(template: Path, data: dict, output: Path) -> Presentation:
    presentation = Presentation(str(template))
    fields = subject.collect_fields(presentation)
    plan = subject.resolve_plan(fields, data, quiet())
    assert plan.missing == []
    subject.fill_presentation(plan, data, quiet())
    presentation.save(str(output))
    return Presentation(str(output))


def shape_named(presentation: Presentation, name: str):
    for slide in presentation.slides:
        for shape in subject.iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False) and shape.name == name:
                return shape
    raise AssertionError(f"shape {name} not found")


def paragraph_marker(paragraph, tag: str) -> bool:
    properties = paragraph._p.find(qn("a:pPr"))
    return properties is not None and properties.find(qn(tag)) is not None


# --- YAML loading -------------------------------------------------------


def test_bom_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text("CLIENT_NAME: Example\n", encoding="utf-8-sig")
    assert subject.load_yaml_data(path) == {"CLIENT_NAME": "Example"}


def test_duplicate_keys_are_rejected(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text("A: one\nA: two\n", encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_yaml_data(path)
    assert "duplicate" in str(info.value).lower()


def test_boolean_like_id_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text("NO: Norway\n", encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_yaml_data(path)
    assert "quote" in str(info.value).lower()


def test_empty_yaml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text("", encoding="utf-8")
    with pytest.raises(subject.InputError):
        subject.load_yaml_data(path)


def test_comment_truncation_is_caught(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text('COLOURED: "<font color="#C00000">red</font>"\n', encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_yaml_data(path)
    assert "COMMENT ERROR" in str(info.value)
    assert "COLOURED: '<font color=\"#C00000\">red</font>'" in str(info.value)


def test_ordinary_trailing_comment_is_allowed(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text("CLIENT_NAME: Example  # set by the deal team\n", encoding="utf-8")
    assert subject.load_yaml_data(path) == {"CLIENT_NAME": "Example"}


def test_escaped_quotes_are_not_mistaken_for_comments(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text('TITLE: "<font color=\\"#C00000\\">red</font>"\n', encoding="utf-8")
    assert subject.load_yaml_data(path) == {"TITLE": '<font color="#C00000">red</font>'}


def test_syntax_error_reports_line_and_fix(tmp_path: Path) -> None:
    path = tmp_path / "data.yaml"
    path.write_text('SIZED: "<font size="14">text</font>"\n', encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_yaml_data(path)
    message = str(info.value)
    assert "Line:" in message and "How to fix it:" in message


# --- value coercion -----------------------------------------------------


@pytest.mark.parametrize("value", [False, "false", "no", "off", 0])
def test_coerce_bool_reads_common_false_values(value) -> None:
    assert subject.coerce_bool(value, "bold", "test") is False


@pytest.mark.parametrize("value", [True, "true", "yes", "on", 1])
def test_coerce_bool_reads_common_true_values(value) -> None:
    assert subject.coerce_bool(value, "bold", "test") is True


def test_coerce_bool_rejects_nonsense() -> None:
    with pytest.raises(ValueError):
        subject.coerce_bool("maybe", "bold", "test")


def test_normalise_colour_accepts_hash_and_plain() -> None:
    assert str(subject.normalise_colour("#C00000")) == "C00000"
    assert str(subject.normalise_colour("c00000")) == "C00000"


def test_normalise_colour_rejects_names() -> None:
    with pytest.raises(ValueError) as info:
        subject.normalise_colour("red", "YAML 'TITLE'")
    assert "invalid colour" in str(info.value)


# --- content definitions ------------------------------------------------


def test_string_content_parses_html() -> None:
    paragraphs = subject.parse_html_content("<b>Bold</b> and <u>under</u>")
    runs = paragraphs[0]["runs"]
    assert runs[0] == {"text": "Bold", "bold": True}
    assert runs[2] == {"text": "under", "underline": True}


def test_literal_entities_stay_text() -> None:
    paragraphs = subject.parse_html_content("&lt;b&gt;not bold&lt;/b&gt;")
    assert paragraphs[0]["runs"] == [{"text": "<b>not bold</b>"}]


def test_nested_span_inherits_style() -> None:
    paragraphs = subject.parse_html_content('<span style="color:#C00000"><b>both</b></span>')
    assert paragraphs[0]["runs"][0]["bold"] is True
    assert paragraphs[0]["runs"][0]["colour"] == "#C00000"


def test_unordered_list_is_bulleted() -> None:
    paragraphs = subject.parse_html_content("<ul><li>one</li></ul>")
    assert paragraphs[0]["bullet"] is True
    assert "number" not in paragraphs[0]


def test_ordered_list_is_numbered() -> None:
    paragraphs = subject.parse_html_content("<ol><li>one</li></ol>")
    assert paragraphs[0]["number"] is True
    assert "bullet" not in paragraphs[0]


def test_nested_list_sets_level() -> None:
    paragraphs = subject.parse_html_content("<ul><li>a<ul><li>b</li></ul></li></ul>")
    assert [paragraph["level"] for paragraph in paragraphs if "level" in paragraph] == [0, 1]


def test_unsupported_tag_lists_supported_tags() -> None:
    with pytest.raises(ValueError) as info:
        subject.parse_html_content("<table></table>")
    assert "Supported tags" in str(info.value)


def test_mismatched_tag_is_reported() -> None:
    with pytest.raises(ValueError) as info:
        subject.parse_html_content("<b>text</u>")
    assert "mismatched" in str(info.value).lower()


def test_unknown_content_key_suggests_correction() -> None:
    with pytest.raises(ValueError) as info:
        subject.value_to_paragraph_definitions({"text": "a", "colourr": "#C00000"}, "YAML 'TITLE'")
    assert "did you mean 'colour'" in str(info.value)


def test_number_and_bullet_conflict_is_rejected(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = {"SUMMARY": {"paragraphs": [{"runs": [{"text": "a"}], "bullet": True, "number": True}]}}
    presentation = Presentation(str(template))
    plan = subject.resolve_plan(subject.collect_fields(presentation), data, quiet())
    with pytest.raises(ValueError) as info:
        subject.fill_presentation(plan, data, quiet())
    assert "not both" in str(info.value)
    assert "YAML 'SUMMARY'" in str(info.value)


def test_content_keys_allow_paragraph_settings() -> None:
    definitions = subject.value_to_paragraph_definitions(
        {"text": "centred", "align": "center", "bullet": False}, "YAML 'DISCLAIMER'"
    )
    assert definitions[0]["align"] == "center"
    assert definitions[0]["runs"] == [{"text": "centred"}]


def test_scalar_content_becomes_text() -> None:
    assert subject.value_to_paragraph_definitions(True, "field")[0]["runs"][0]["text"] == "true"
    assert subject.value_to_paragraph_definitions(42, "field")[0]["runs"][0]["text"] == "42"


# --- template matching --------------------------------------------------


def test_token_match_tolerates_spaces(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    presentation = Presentation(str(template))
    fields = subject.collect_fields(presentation)
    plan = subject.resolve_plan(fields, {"SUMMARY": "filled"}, quiet())
    assert plan.missing == []
    assert plan.matches["SUMMARY"][0].matched_by == "token"


def test_missing_id_suggests_close_match(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    fields = subject.collect_fields(Presentation(str(template)))
    message = subject.build_missing_message(["CLIENT_NAM"], fields)
    assert "did you mean 'CLIENT_NAME'" in message


def test_duplicate_target_warns(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    fields = subject.collect_fields(Presentation(str(template)))
    plan = subject.resolve_plan(fields, {"CLIENT_NAME": "a", "client_name": "b"}, quiet())
    assert plan.matches["CLIENT_NAME"][0] is plan.matches["client_name"][0]


# --- writing ------------------------------------------------------------


def test_fill_named_and_token_shapes(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = {
        "CLIENT_NAME": "<b>Example Energy</b>",
        "SUMMARY": {"paragraphs": [{"bullet": True, "runs": [{"text": "first"}]}, {"runs": [{"text": "second"}]}]},
    }
    presentation = fill_and_reopen(template, data, tmp_path / "out.pptx")

    client = shape_named(presentation, "CLIENT_NAME")
    assert client.text_frame.text == "Example Energy"
    assert client.text_frame.paragraphs[0].runs[0].font.bold is True

    summary = [shape for shape in subject.collect_fields(presentation) if "first" in shape.text][0]
    assert paragraph_marker(summary.shape.text_frame.paragraphs[0], "a:buChar")
    assert summary.shape.text_frame.paragraphs[1].text == "second"


def test_ordered_list_from_yaml_writes_numbering(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = {"SUMMARY": "<ol><li>one</li><li>two</li></ol>"}
    presentation = fill_and_reopen(template, data, tmp_path / "out.pptx")
    summary = [shape for shape in subject.collect_fields(presentation) if "one" in shape.text][0]
    assert paragraph_marker(summary.shape.text_frame.paragraphs[0], "a:buAutoNum")


def test_template_file_is_not_modified(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    before = template.read_bytes()
    fill_and_reopen(template, {"CLIENT_NAME": "changed"}, tmp_path / "out.pptx")
    assert template.read_bytes() == before


# --- output paths -------------------------------------------------------


def make_job(tmp_path: Path, **overrides) -> subject.Job:
    job = subject.Job(
        template=tmp_path / "Template.pptx",
        data=tmp_path / "Deal data.yaml",
        output_dir=tmp_path,
    )
    for key, value in overrides.items():
        setattr(job, key, value)
    return job


def test_second_run_does_not_clobber(tmp_path: Path) -> None:
    first_pptx, _ = subject.resolve_output_paths(make_job(tmp_path), quiet())
    first_pptx.write_text("x", encoding="utf-8")
    second_pptx, _ = subject.resolve_output_paths(make_job(tmp_path), quiet())
    assert second_pptx != first_pptx
    assert second_pptx.stem.endswith("(2)")


def test_overwrite_flag_reuses_the_name(tmp_path: Path) -> None:
    first_pptx, _ = subject.resolve_output_paths(make_job(tmp_path), quiet())
    first_pptx.write_text("x", encoding="utf-8")
    second_pptx, _ = subject.resolve_output_paths(make_job(tmp_path, overwrite=True), quiet())
    assert second_pptx == first_pptx


def test_output_name_contains_names_and_date(tmp_path: Path) -> None:
    pptx_path, pdf_path = subject.resolve_output_paths(make_job(tmp_path), quiet())
    assert pptx_path.name.startswith("Template - Deal data - ")
    assert pdf_path.suffix == ".pdf"


# --- command line -------------------------------------------------------


def write_data(tmp_path: Path, text: str) -> Path:
    path = tmp_path / "data.yaml"
    path.write_text(text, encoding="utf-8")
    return path


def test_no_pdf_writes_only_pptx(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    code = subject.main([str(template), str(data), "-o", str(output), "--no-pdf", "--quiet"])
    assert code == 0
    assert [path.suffix for path in output.iterdir()] == [".pptx"]


def test_dry_run_writes_nothing(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    code = subject.main([str(template), str(data), "-o", str(output), "--no-pdf", "--dry-run", "-q"])
    assert code == 0
    assert not output.exists()


def test_missing_id_exits_three(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "NOPE: Example\n")
    code = subject.main([str(template), str(data), "--no-pdf", "-q"])
    assert code == 3


def test_bad_yaml_exits_two(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, 'A: "<b size="2">x</b>"\n')
    code = subject.main([str(template), str(data), "--no-pdf", "-q"])
    assert code == 2


def test_list_fields_prints_shape_names(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    code = subject.main([str(template), "--list-fields"])
    assert code == 0
    assert "CLIENT_NAME" in capsys.readouterr().out


def test_missing_data_argument_exits_two(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    assert subject.main([str(template), "--no-pdf", "-q"]) == 2


def test_export_failure_exits_four_and_keeps_pptx(tmp_path: Path, monkeypatch) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    output.mkdir()

    def fail(*args, **kwargs):
        raise subject.ExportError("PDF export failed: boom")

    monkeypatch.setattr(subject, "export_pptx_to_pdf", fail)
    monkeypatch.setattr(subject, "build_job", lambda arguments: subject.Job(
        template=template, data=data, output_dir=output, export_pdf=True
    ))
    code = subject.main([str(template), str(data), "--quiet"])
    assert code == 4
    assert [path.suffix for path in output.iterdir()] == [".pptx"]
