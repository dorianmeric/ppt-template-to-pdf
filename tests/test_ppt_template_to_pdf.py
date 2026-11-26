"""Tests for ppt_template_to_pdf.

Only the pure functions and the PPTX path are covered. PDF export needs a real
PowerPoint install and is exercised manually.
"""

from __future__ import annotations

import importlib.util
import re
import sys
import tomllib
import zipfile
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


# --- JSON loading -------------------------------------------------------


def test_json_data_is_loaded(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('{"CLIENT_NAME": "Example", "SUMMARY": {"text": "hi"}}', encoding="utf-8")
    assert subject.load_data_file(path) == {"CLIENT_NAME": "Example", "SUMMARY": {"text": "hi"}}


def test_json_bom_is_accepted(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_bytes('{"CLIENT_NAME": "Example"}'.encode("utf-8-sig"))
    assert subject.load_json_data(path) == {"CLIENT_NAME": "Example"}


def test_json_keeps_untouched_number_and_bool(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('{"SIZE": 9, "BOLD": true}', encoding="utf-8")
    assert subject.load_json_data(path) == {"SIZE": 9, "BOLD": True}


def test_json_duplicate_keys_are_rejected(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('{"A": "one", "A": "two"}', encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_json_data(path)
    message = str(info.value)
    assert "DUPLICATE KEY" in message and "'A'" in message


def test_json_trailing_comma_reports_line_and_fix(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('{\n  "CLIENT_NAME": "Example",\n}\n', encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_json_data(path)
    message = str(info.value)
    assert "JSON SYNTAX ERROR" in message
    assert "Line:" in message and "How to fix it:" in message
    assert "trailing comma" in message


def test_json_comment_is_explained(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('{\n  "A": "one"\n  // set by the deal team\n}', encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_json_data(path)
    assert "no comments" in str(info.value)


def test_json_empty_file_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text("   \n", encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_json_data(path)
    assert "empty" in str(info.value).lower()


def test_json_top_level_list_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('["A", "B"]', encoding="utf-8")
    with pytest.raises(subject.InputError) as info:
        subject.load_json_data(path)
    assert "top level" in str(info.value)


def test_json_html_value_is_parsed_as_formatting(tmp_path: Path) -> None:
    path = tmp_path / "data.json"
    path.write_text('{"TITLE": "<font color=\\"#C00000\\">red</font>"}', encoding="utf-8")
    assert subject.load_json_data(path) == {"TITLE": '<font color="#C00000">red</font>'}
    paragraphs = subject.value_to_paragraph_definitions("<b>bold</b>", "JSON 'TITLE'")
    assert paragraphs[0]["runs"][0]["bold"] is True


# --- dependency guard ---------------------------------------------------

REQUIREMENTS_BLOCK = re.compile(r"^# /// script\n(?P<body>(?:^#.*\n)*?)^# ///$", re.MULTILINE)


def third_party_imports(source: str) -> set[str]:
    """Top-level module names imported anywhere in a script, minus the standard library."""
    import ast

    names: set[str] = set()
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.Import):
            names.update(alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            names.add(node.module.split(".")[0])
    return {name for name in names if name not in sys.stdlib_module_names and name != "__future__"}


def test_every_third_party_import_is_mapped_to_a_package() -> None:
    """A new import must not sneak in without a PyPI name for the guard."""
    source = Path("ppt_template_to_pdf.py").read_text(encoding="utf-8")
    mapped = set(subject.REQUIRED_MODULES) | set(subject.WINDOWS_REQUIRED_MODULES)
    assert third_party_imports(source) == mapped


def test_mapping_uses_the_real_pypi_names() -> None:
    assert subject.REQUIRED_MODULES["pptx"].startswith("python-pptx")
    assert subject.REQUIRED_MODULES["yaml"].startswith("PyYAML")
    assert subject.WINDOWS_REQUIRED_MODULES["win32com"].startswith("pywin32")


def test_required_modules_follow_the_platform(monkeypatch) -> None:
    monkeypatch.setattr(subject.sys, "platform", "linux")
    assert "win32com" not in subject.required_modules()
    monkeypatch.setattr(subject.sys, "platform", "win32")
    assert "win32com" in subject.required_modules()


def test_pep723_block_declares_the_same_dependencies() -> None:
    source = Path("ppt_template_to_pdf.py").read_text(encoding="utf-8")
    match = REQUIREMENTS_BLOCK.search(source)
    assert match, "the PEP 723 block is missing from the top of the script"
    # The block is TOML once the comment markers are stripped, so parse it as TOML
    # rather than comparing strings.
    body = "\n".join(
        line[2:] if line.startswith("# ") else line.lstrip("#")
        for line in match.group("body").splitlines()
    )
    metadata = tomllib.loads(body)
    assert metadata["requires-python"] == ">=3.10"
    dependencies = metadata["dependencies"]

    # pyproject.toml is what pip and uv use for an installed copy; the inline block
    # must not drift away from it.
    project = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))
    assert set(dependencies) == set(project["project"]["dependencies"])

    for requirement in subject.REQUIRED_MODULES.values():
        assert requirement in dependencies, f"{requirement} is missing from the PEP 723 block"
    for requirement in subject.WINDOWS_REQUIRED_MODULES.values():
        assert any(dep.startswith(requirement) for dep in dependencies), (
            f"{requirement} is missing from the PEP 723 block"
        )


def test_guard_is_skipped_in_a_frozen_build(monkeypatch) -> None:
    monkeypatch.setattr(subject, "running_frozen", lambda: True)
    monkeypatch.setattr(
        subject.importlib.util, "find_spec", lambda name: pytest.fail("probed a module")
    )
    subject.ensure_dependencies()


def test_guard_says_nothing_when_everything_is_installed(monkeypatch, capsys) -> None:
    monkeypatch.setattr(subject, "running_frozen", lambda: False)
    monkeypatch.setattr(subject.importlib.util, "find_spec", lambda name: object())
    subject.ensure_dependencies()
    captured = capsys.readouterr()
    assert captured.out == "" and captured.err == ""


def test_guard_exits_with_the_command_when_not_interactive(monkeypatch, capsys) -> None:
    monkeypatch.setattr(subject, "running_frozen", lambda: False)
    monkeypatch.setattr(subject.importlib.util, "find_spec", lambda name: None)
    monkeypatch.setattr(subject.sys.stdin, "isatty", lambda: False, raising=False)

    with pytest.raises(SystemExit) as info:
        subject.ensure_dependencies()
    assert info.value.code == 1

    errors = capsys.readouterr().err
    assert "python-pptx" in errors and "PyYAML" in errors
    requirements = subject.required_modules()
    packages = " ".join(f'"{requirements[module]}"' for module in sorted(requirements))
    expected = f'"{subject.sys.executable}" -m pip install {packages}'
    assert expected in errors
    assert "nothing was installed" in errors.lower()


def test_guard_installs_when_the_user_says_yes(monkeypatch, capsys) -> None:
    calls: list[list[str]] = []
    prompts: list[str] = []

    def fake_input(prompt: str = "") -> str:
        prompts.append(prompt)
        return "y"

    monkeypatch.setattr(subject, "running_frozen", lambda: False)
    monkeypatch.setattr(subject.importlib.util, "find_spec", lambda name: None)
    monkeypatch.setattr(subject.sys.stdin, "isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input", fake_input)
    monkeypatch.setattr(subject.importlib, "invalidate_caches", lambda: calls.append(["invalidate"]))
    monkeypatch.setattr(subject.subprocess, "run", lambda command, check: calls.append(command))

    subject.ensure_dependencies()

    assert prompts == [f"Install now into {subject.sys.executable}? [y/N]: "]
    assert calls[0][1:4] == ["-m", "pip", "install"]
    assert "python-pptx>=1.0.2" in calls[0]
    assert calls[-1] == ["invalidate"]
    assert "Installing" in capsys.readouterr().out


def test_guard_respects_no(monkeypatch, capsys) -> None:
    monkeypatch.setattr(subject, "running_frozen", lambda: False)
    monkeypatch.setattr(subject.importlib.util, "find_spec", lambda name: None)
    monkeypatch.setattr(subject.sys.stdin, "isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input", lambda prompt="": "n")
    monkeypatch.setattr(
        subject.subprocess, "run", lambda *a, **k: pytest.fail("installed anyway")
    )
    with pytest.raises(SystemExit) as info:
        subject.ensure_dependencies()
    assert info.value.code == 1
    assert "Nothing was installed" in capsys.readouterr().err


def test_guard_survives_eof(monkeypatch, capsys) -> None:
    monkeypatch.setattr(subject, "running_frozen", lambda: False)
    monkeypatch.setattr(subject.importlib.util, "find_spec", lambda name: None)
    monkeypatch.setattr(subject.sys.stdin, "isatty", lambda: True, raising=False)

    def raise_eof(prompt=""):
        raise EOFError

    monkeypatch.setattr("builtins.input", raise_eof)
    with pytest.raises(SystemExit) as info:
        subject.ensure_dependencies()
    assert info.value.code == 1


def test_guard_reports_a_failed_pip_install(monkeypatch, capsys) -> None:
    def fail(command, check):
        raise subject.subprocess.CalledProcessError(2, command)

    monkeypatch.setattr(subject, "running_frozen", lambda: False)
    monkeypatch.setattr(subject.importlib.util, "find_spec", lambda name: None)
    monkeypatch.setattr(subject.sys.stdin, "isatty", lambda: True, raising=False)
    monkeypatch.setattr("builtins.input", lambda prompt="": "yes")
    monkeypatch.setattr(subject.subprocess, "run", fail)

    with pytest.raises(SystemExit) as info:
        subject.ensure_dependencies()
    assert info.value.code == 1
    assert "pip install failed with exit code 2" in capsys.readouterr().err


# --- document properties ------------------------------------------------


def test_properties_block_is_split_from_fields() -> None:
    shape_data, properties = subject.split_document_properties(
        {"TITLE": "Slide title", "PROPERTIES": {"title": "Deck title"}}
    )
    assert shape_data == {"TITLE": "Slide title"}
    assert properties == {"title": "Deck title"}


def test_properties_key_folds_case() -> None:
    shape_data, properties = subject.split_document_properties(
        {"Properties": {"Subject": "Quarterly review"}, "TITLE": "x"}
    )
    assert shape_data == {"TITLE": "x"}
    assert properties == {"Subject": "Quarterly review"}


def test_properties_block_must_be_a_mapping() -> None:
    with pytest.raises(subject.InputError) as info:
        subject.split_document_properties({"PROPERTIES": "title: Q3"})
    assert "mapping" in str(info.value)


@pytest.mark.parametrize(
    "written,expected",
    [
        ("title", "title"),
        ("Title", "title"),
        ("SUBJECT", "subject"),
        ("last modified by", "last_modified_by"),
        ("lastModifiedBy", "last_modified_by"),
        ("Last-Modified-By", "last_modified_by"),
        ("contentStatus", "content_status"),
        ("LastPrinted", "last_printed"),
    ],
)
def test_property_names_are_normalised(written: str, expected: str) -> None:
    assert subject.normalise_property_name(written) == expected


def test_property_values_are_converted() -> None:
    properties = subject.normalise_document_properties(
        {
            "title": "Q3 update",
            "revision": "7",
            "keywords": ["q3", "board", " "],
            "created": "2026-09-30T09:00:00",
            "version": 1.2,
        }
    )
    assert properties["title"] == "Q3 update"
    assert properties["revision"] == 7
    assert properties["keywords"] == "q3, board"
    assert properties["version"] == "1.2"
    assert properties["created"].isoformat() == "2026-09-30T09:00:00"


def test_unknown_property_is_rejected_with_a_suggestion() -> None:
    with pytest.raises(subject.InputError) as info:
        subject.normalise_document_properties({"titel": "Q3 update"})
    message = str(info.value)
    assert "unknown document property" in message
    assert "did you mean 'title'" in message


def test_bad_property_dates_are_rejected() -> None:
    with pytest.raises(subject.InputError) as info:
        subject.normalise_document_properties({"created": "last Tuesday"})
    assert "ISO 8601" in str(info.value)


def test_bad_revision_is_rejected() -> None:
    with pytest.raises(subject.InputError):
        subject.normalise_document_properties({"revision": "first"})


def test_properties_are_written_to_the_pptx(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(
        tmp_path,
        "CLIENT_NAME: Example\n"
        "PROPERTIES:\n"
        "  title: Q3 Results\n"
        "  subject: Quarterly review\n"
        "  author: Corporate Finance\n"
        "  keywords: [q3, board pack]\n"
        "  revision: 3\n"
        "  created: 2026-09-30T09:00:00\n",
    )
    output = tmp_path / "out"
    args = [str(template), str(data), "-o", str(output), "--no-pdf", "-q"]
    assert subject.main(args) == 0

    written = Presentation(str(next(output.glob("*.pptx"))))
    assert written.core_properties.title == "Q3 Results"
    assert written.core_properties.subject == "Quarterly review"
    assert written.core_properties.author == "Corporate Finance"
    assert written.core_properties.keywords == "q3, board pack"
    assert written.core_properties.revision == 3
    assert written.core_properties.created.isoformat() == "2026-09-30T09:00:00"


def test_properties_only_file_writes_no_shape(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_json_data(tmp_path, '{"PROPERTIES": {"title": "Only metadata"}}')
    output = tmp_path / "out"
    args = [str(template), str(data), "-o", str(output), "--no-pdf", "-q"]
    assert subject.main(args) == 0

    written = Presentation(str(next(output.glob("*.pptx"))))
    assert written.core_properties.title == "Only metadata"
    assert shape_named(written, "CLIENT_NAME").text_frame.text == "old client"


def test_unknown_property_exits_two(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "PROPERTIES:\n  titel: Q3\n")
    assert subject.main([str(template), str(data), "--no-pdf", "-q"]) == 2


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


def test_collect_fields_covers_slides_layouts_and_masters(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    fields = subject.collect_fields(Presentation(str(template)))
    regions = {field.region for field in fields}
    assert regions == {"slide", "layout", "master"}
    assert any(field.where.startswith("slide ") for field in fields)
    assert any(field.where.startswith("layout '") for field in fields)
    assert any(field.where.startswith("master '") for field in fields)
    assert len({id(field.shape) for field in fields}) == len(fields)


def test_field_on_the_slide_master_is_filled() -> None:
    presentation = Presentation()
    master_title = presentation.slide_master.placeholders[0]
    master_title.text_frame.text = "{{MASTER_TITLE}}"

    data = {"MASTER_TITLE": "Every slide shows this"}
    plan = subject.resolve_plan(subject.collect_fields(presentation), data, quiet())
    assert plan.missing == []
    assert [field.region for field in plan.matches["MASTER_TITLE"]] == ["master"]

    subject.fill_presentation(plan, data, quiet())
    assert master_title.text_frame.text == "Every slide shows this"
    assert "MASTER_TITLE" not in master_title.text_frame.text


def test_field_on_a_slide_layout_is_filled() -> None:
    presentation = Presentation()
    layout_title = presentation.slide_layouts[1].placeholders[0]
    layout_title.text_frame.text = "{{LAYOUT_TITLE}}"

    data = {"LAYOUT_TITLE": "Layout value"}
    plan = subject.resolve_plan(subject.collect_fields(presentation), data, quiet())
    assert plan.missing == []
    assert plan.matches["LAYOUT_TITLE"][0].region == "layout"

    subject.fill_presentation(plan, data, quiet())
    assert layout_title.text_frame.text == "Layout value"


def test_master_field_written_through_the_cli(tmp_path: Path) -> None:
    template_path = build_template(tmp_path / "Template.pptx")
    presentation = Presentation(str(template_path))
    presentation.slide_layouts[1].placeholders[0].text_frame.text = "{{LAYOUT_TITLE}}"
    presentation.save(str(template_path))

    data = write_data(tmp_path, "LAYOUT_TITLE: From the layout\nCLIENT_NAME: Example\n")
    output = tmp_path / "out"
    args = [str(template_path), str(data), "-o", str(output), "--no-pdf", "-q"]
    assert subject.main(args) == 0

    written = Presentation(str(next(output.glob("*.pptx"))))
    layout_title = written.slide_layouts[1].placeholders[0]
    assert layout_title.text_frame.text == "From the layout"


def test_same_field_repeated_on_every_slide_is_filled(tmp_path: Path) -> None:
    presentation = Presentation()
    for slide_number in range(3):
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        named = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        named.name = "CLIENT_NAME"
        named.text_frame.text = "old client"
        token = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
        token.text_frame.text = "Prepared for {{ CLIENT_NAME }}"
    template = tmp_path / "Template.pptx"
    presentation.save(str(template))

    data = {"CLIENT_NAME": "Example Energy"}
    fields = subject.collect_fields(Presentation(str(template)))
    plan = subject.resolve_plan(fields, data, quiet())
    assert plan.missing == []
    assert len(plan.matches["CLIENT_NAME"]) == 6
    assert plan.field_count == 6

    filled = fill_and_reopen(template, data, tmp_path / "out.pptx")
    filled_fields = subject.collect_fields(filled)
    named_shapes = [field for field in filled_fields if field.name == "CLIENT_NAME"]
    assert [field.text for field in named_shapes] == ["Example Energy"] * 3
    # A token match replaces the whole text frame, so all six boxes end up equal.
    token_shapes = [field for field in filled_fields if field.name.startswith("TextBox")]
    assert [field.text for field in token_shapes] == ["Example Energy"] * 3
    assert not any("{{" in field.text for field in filled_fields)


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


def test_resolve_template_accepts_legacy_suffixes(tmp_path: Path) -> None:
    for name in ("Old.ppt", "Old.pot"):
        path = tmp_path / name
        path.write_bytes(b"stub")
        assert subject.resolve_template(path) == path.resolve()


def test_modern_template_is_used_directly(tmp_path: Path, monkeypatch) -> None:
    template = build_template(tmp_path / "Template.pptx")
    calls = []
    monkeypatch.setattr(subject, "convert_with_powerpoint", lambda *a, **k: calls.append(a))
    with subject.usable_template(template, quiet()) as template_path:
        assert template_path == template
    assert calls == []


def test_legacy_template_is_converted_and_cleaned_up(tmp_path: Path, monkeypatch) -> None:
    legacy = tmp_path / "Template.ppt"
    legacy.write_bytes(b"binary stub")
    real_deck = build_template(tmp_path / "real.pptx")
    seen = {}

    def fake_convert(source, target, file_format, action, error_type, hint):
        seen["source"] = source
        seen["format"] = file_format
        target.write_bytes(real_deck.read_bytes())

    monkeypatch.setattr(subject, "convert_with_powerpoint", fake_convert)
    monkeypatch.setattr(subject, "pywin32_available", lambda: True)
    monkeypatch.setattr(subject.sys, "platform", "win32")
    with subject.usable_template(legacy, quiet()) as template_path:
        assert template_path != legacy
        assert template_path.suffix == ".pptx"
        assert Presentation(str(template_path)).slides
    assert seen["source"] == legacy
    assert seen["format"] == subject.PP_SAVE_AS_OPEN_XML_PRESENTATION
    assert not template_path.exists()


def test_legacy_template_without_powerpoint_explains_the_fix(tmp_path: Path, monkeypatch) -> None:
    legacy = tmp_path / "Template.ppt"
    legacy.write_bytes(b"binary stub")
    monkeypatch.setattr(subject, "pywin32_available", lambda: False)
    with pytest.raises(subject.InputError) as info:
        with subject.usable_template(legacy, quiet()):
            pass
    message = str(info.value)
    assert "Save As" in message
    assert ".pptx" in message


def test_legacy_template_without_powerpoint_exits_two(tmp_path: Path, monkeypatch) -> None:
    legacy = tmp_path / "Template.ppt"
    legacy.write_bytes(b"binary stub")
    data = write_data(tmp_path, "A: b\n")
    monkeypatch.setattr(subject, "pywin32_available", lambda: False)
    assert subject.main([str(legacy), str(data), "--no-pdf", "-q"]) == 2


def test_output_name_contains_names_and_date(tmp_path: Path) -> None:
    pptx_path, pdf_path = subject.resolve_output_paths(make_job(tmp_path), quiet())
    assert pptx_path.name.startswith("Template - Deal data - ")
    assert pdf_path.suffix == ".pdf"


def test_custom_name_names_both_files(tmp_path: Path) -> None:
    pptx_path, pdf_path = subject.resolve_output_paths(make_job(tmp_path, output_name="Q3 update"), quiet())
    assert pptx_path.name == "Q3 update.pptx"
    assert pdf_path.name == "Q3 update.pdf"


def test_custom_name_keeps_the_numbering(tmp_path: Path) -> None:
    job = make_job(tmp_path, output_name="Q3 update")
    first_pptx, _ = subject.resolve_output_paths(job, quiet())
    first_pptx.write_bytes(b"")
    second_pptx, second_pdf = subject.resolve_output_paths(job, quiet())
    assert second_pptx.name == "Q3 update (2).pptx"
    assert second_pdf.name == "Q3 update (2).pdf"


def test_custom_name_with_no_pdf_still_returns_both_paths(tmp_path: Path) -> None:
    pptx_path, pdf_path = subject.resolve_output_paths(
        make_job(tmp_path, output_name="Q3 update", export_pdf=False), quiet()
    )
    assert pptx_path.name == "Q3 update.pptx"
    assert pdf_path.name == "Q3 update.pdf"


@pytest.mark.parametrize("typed", ["Report", "Report.pdf", "Report.pptx", "Report.PDF", "  Report  "])
def test_name_drops_a_typed_extension(typed: str) -> None:
    assert subject.normalise_output_name(typed) == "Report"


@pytest.mark.parametrize(
    "bad",
    ["", "   ", "a/b", "a\\b", "a:b", "a*b", "a?b", "a|b", "a\"b", "..hidden", "trailing.", "CON", "lpt1", "nul"],
)
def test_bad_names_are_rejected(bad: str) -> None:
    with pytest.raises(subject.InputError):
        subject.normalise_output_name(bad)


# --- command line -------------------------------------------------------


def write_data(tmp_path: Path, text: str) -> Path:
    path = tmp_path / "data.yaml"
    path.write_text(text, encoding="utf-8")
    return path


def test_progress_is_printed_by_default(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    assert subject.main([str(template), str(data), "-o", str(output), "--no-pdf"]) == 0

    printed = capsys.readouterr().out
    assert "Reading data.yaml" in printed
    assert "Opening Template.pptx" in printed
    assert "Searching" in printed
    assert "Matched 1 ID to 1 text shape" in printed
    assert "Writing 1 value into 1 text shape" in printed
    assert "Created PPTX:" in printed


def test_quiet_run_prints_nothing_on_success(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    assert subject.main([str(template), str(data), "-o", str(output), "--no-pdf", "-q"]) == 0
    assert capsys.readouterr().out == ""


def test_verbose_run_lists_each_write(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    args = [str(template), str(data), "-o", str(output), "--no-pdf", "-v"]
    assert subject.main(args) == 0

    printed = capsys.readouterr().out
    assert "left unchanged: slide 1" in printed
    assert "CLIENT_NAME -> slide 1 'CLIENT_NAME' (shape name)" in printed
    assert "PDF export skipped" in printed


def test_dry_run_reports_the_properties_it_would_set(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(
        tmp_path,
        "CLIENT_NAME: Example\nPROPERTIES:\n  title: Q3\n  author: Deal team\n",
    )
    assert subject.main([str(template), str(data), "--no-pdf", "--dry-run"]) == 0

    printed = capsys.readouterr().out
    assert "Dry run: 1 ID matched 1 text shape" in printed
    assert "2 document properties would be set" in printed
    assert "Nothing was written." in printed


# --- release zip --------------------------------------------------------


def load_release_tool():
    """Import tools/make_github_release.py, which is not on the pytest path."""
    path = Path(__file__).resolve().parent.parent / "tools" / "make_github_release.py"
    spec = importlib.util.spec_from_file_location("make_github_release", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def project_version(tool) -> str:
    return tool.read_version(
        tool.PROJECT_ROOT / "pyproject.toml", tool.PROJECT_VERSION_PATTERN, "version"
    )


def test_release_zip_holds_only_the_release_files(tmp_path: Path) -> None:
    tool = load_release_tool()
    version = project_version(tool)
    archive_path = tool.build_archive(version, tmp_path)

    with zipfile.ZipFile(archive_path) as archive:
        names = archive.namelist()
    assert names == [f"ppt-template-to-pdf-{version}/{name}" for name in tool.RELEASE_FILES]
    assert not any(
        "/tests/" in name or "/assets/" in name or name.endswith("requirements-dev.txt")
        for name in names
    )


def test_release_zip_versions_must_agree(tmp_path: Path) -> None:
    tool = load_release_tool()
    with pytest.raises(SystemExit) as info:
        tool.main(["--expect-version", "9.9.9", "--output-dir", str(tmp_path)])
    assert "Bump pyproject.toml" in str(info.value)


def test_release_zip_builds_for_the_current_version(tmp_path: Path, capsys) -> None:
    tool = load_release_tool()
    version = project_version(tool)
    arguments = ["--zip-only", "--expect-version", version, "--output-dir", str(tmp_path)]
    assert tool.main(arguments) == 0
    assert (tmp_path / f"ppt-template-to-pdf-{version}.zip").is_file()
    assert f"contains ppt-template-to-pdf-{version}/README.md" in capsys.readouterr().out


def test_host_target_is_a_known_pair() -> None:
    tool = load_release_tool()
    os_slug, arch = tool.host_target()
    assert os_slug in {"windows", "macos", "linux"}
    assert arch in {"x64", "x86", "arm64", "x86_64", "aarch64"}


@pytest.mark.parametrize(
    "os_slug,arch,expected",
    [
        ("windows", "x64", "ppt-template-to-pdf-1.9.1-windows-x64.exe"),
        ("windows", "x86", "ppt-template-to-pdf-1.9.1-windows-x86.exe"),
        ("macos", "arm64", "ppt-template-to-pdf-1.9.1-macos-arm64"),
        ("macos", "x86_64", "ppt-template-to-pdf-1.9.1-macos-x86_64"),
        ("linux", "x86_64", "ppt-template-to-pdf-1.9.1-linux-x86_64"),
    ],
)
def test_binary_names_match_the_published_assets(os_slug: str, arch: str, expected: str) -> None:
    tool = load_release_tool()
    assert tool.release_binary_name("1.9.1", os_slug, arch) == expected


def test_nuitka_command_mirrors_the_workflow(tmp_path: Path) -> None:
    tool = load_release_tool()

    windows = tool.nuitka_command("1.9.1", "windows", "x64", tmp_path)
    assert windows[:3] == [sys.executable, "-m", "nuitka"]
    assert "--onefile" in windows
    assert "--output-filename=ppt-template-to-pdf-1.9.1-windows-x64.exe" in windows
    assert "--windows-console-mode=force" in windows
    assert "--include-package=win32com" in windows
    assert "--file-version=1.9.1" in windows
    assert windows[-1].endswith("ppt_template_to_pdf.py")

    linux = tool.nuitka_command("1.9.1", "linux", "x86_64", tmp_path)
    assert "--output-filename=ppt-template-to-pdf-1.9.1-linux-x86_64" in linux
    assert not any("win32com" in part or "windows-" in part for part in linux)


def test_binary_build_runs_nuitka_then_the_binary(tmp_path: Path, monkeypatch) -> None:
    tool = load_release_tool()
    binary = tmp_path / tool.release_binary_name("1.9.1", *tool.host_target())
    calls: list[list[str]] = []

    class Result:
        def __init__(self, stdout: str = "") -> None:
            self.returncode = 0
            self.stdout = stdout
            self.stderr = ""

    def fake_run(command, **kwargs):
        calls.append(list(command))
        if "nuitka" in command:
            binary.write_bytes(b"fake binary")
            return Result()
        return Result(stdout="ppt_template_to_pdf 1.9.1\n")

    monkeypatch.setattr(tool.subprocess, "run", fake_run)
    monkeypatch.setattr(tool, "find_spec", lambda name: object())

    built = tool.build_binary("1.9.1", tmp_path)
    assert built == binary
    assert calls[0][:3] == [sys.executable, "-m", "nuitka"]
    assert calls[1][0] == str(binary)


def test_binary_build_fails_when_the_binary_is_wrong(tmp_path: Path, monkeypatch) -> None:
    tool = load_release_tool()
    binary = tmp_path / tool.release_binary_name("1.9.1", *tool.host_target())

    class Result:
        def __init__(self, stdout: str = "") -> None:
            self.returncode = 0
            self.stdout = stdout
            self.stderr = ""

    def fake_run(command, **kwargs):
        if "nuitka" in command:
            binary.write_bytes(b"fake binary")
            return Result()
        return Result(stdout="ppt_template_to_pdf 0.0.0\n")

    monkeypatch.setattr(tool.subprocess, "run", fake_run)
    monkeypatch.setattr(tool, "find_spec", lambda name: object())
    with pytest.raises(SystemExit) as info:
        tool.build_binary("1.9.1", tmp_path)
    assert "did not report version" in str(info.value)


def test_zip_only_skips_the_binary(tmp_path: Path, monkeypatch) -> None:
    tool = load_release_tool()
    monkeypatch.setattr(tool, "build_binary", lambda *a, **k: pytest.fail("built a binary"))
    version = project_version(tool)
    assert tool.main(["--zip-only", "--output-dir", str(tmp_path)]) == 0
    assert (tmp_path / f"ppt-template-to-pdf-{version}.zip").is_file()


def test_binary_only_skips_the_zip(tmp_path: Path, monkeypatch) -> None:
    tool = load_release_tool()

    def fake_build_binary(version, output_dir):
        path = output_dir / "binary"
        path.write_bytes(b"fake")
        return path

    monkeypatch.setattr(tool, "build_archive", lambda *a, **k: pytest.fail("built the zip"))
    monkeypatch.setattr(tool, "build_binary", fake_build_binary)
    assert tool.main(["--binary-only", "--output-dir", str(tmp_path)]) == 0


def test_script_and_pyproject_versions_match() -> None:
    pyproject = Path("pyproject.toml").read_text(encoding="utf-8")
    script = Path("ppt_template_to_pdf.py").read_text(encoding="utf-8")
    declared = re.search(r'^version = "([^"]+)"', pyproject, re.MULTILINE)
    shipped = re.search(r'^__version__ = "([^"]+)"', script, re.MULTILINE)
    assert declared and shipped
    assert declared.group(1) == shipped.group(1) == subject.__version__


def test_reporter_flushes_each_line(monkeypatch) -> None:
    recorded: dict = {}

    def fake_print(*args, **kwargs):
        recorded["args"] = args
        recorded["kwargs"] = kwargs

    monkeypatch.setattr("builtins.print", fake_print)
    subject.Reporter().info("progress")
    assert recorded["args"] == ("progress",)
    assert recorded["kwargs"]["flush"] is True


def test_pywin32_check_uses_a_real_import(monkeypatch) -> None:
    """A frozen build has the modules embedded, so the check must try to import."""
    import builtins

    original_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name.startswith("win32com"):
            raise ImportError(name)
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    assert subject.pywin32_available() is False


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


def write_json_data(tmp_path: Path, text: str) -> Path:
    path = tmp_path / "data.json"
    path.write_text(text, encoding="utf-8")
    return path


def test_json_data_file_writes_the_deck(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_json_data(tmp_path, '{"CLIENT_NAME": "Example", "SUMMARY": "Done"}')
    output = tmp_path / "out"
    code = subject.main([str(template), str(data), "-o", str(output), "--no-pdf", "--quiet"])
    assert code == 0
    written = next(output.glob("*.pptx"))
    assert written.name.startswith("Template - data - ")
    assert shape_named(Presentation(str(written)), "CLIENT_NAME").text_frame.text == "Example"


def test_missing_json_id_exits_three_and_names_json(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_json_data(tmp_path, '{"NOPE": "Example"}')
    code = subject.main([str(template), str(data), "--no-pdf"])
    assert code == 3
    assert "JSON IDs were not found" in capsys.readouterr().err


def test_unsupported_data_suffix_exits_two(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = tmp_path / "data.txt"
    data.write_text("{}", encoding="utf-8")
    assert subject.main([str(template), str(data), "--no-pdf", "-q"]) == 2


def test_list_fields_prints_shape_names(tmp_path: Path, capsys) -> None:
    template = build_template(tmp_path / "Template.pptx")
    code = subject.main([str(template), "--list-fields"])
    assert code == 0
    output = capsys.readouterr().out
    assert "CLIENT_NAME" in output
    assert "Where" in output
    assert "on layouts and masters" in output


def test_list_fields_shows_a_token_on_the_master(capsys) -> None:
    presentation = Presentation()
    presentation.slide_master.placeholders[0].text_frame.text = "{{FOOTER_TEXT}}"
    subject.print_field_list(subject.collect_fields(presentation))
    output = capsys.readouterr().out
    assert "master" in output
    assert "token FOOTER_TEXT" in output


def test_name_flag_names_the_files(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    output = tmp_path / "out"
    args = [str(template), str(data), "-o", str(output), "--no-pdf", "-q", "-n", "Q3 update.pdf"]
    assert subject.main(args) == 0
    assert [path.name for path in output.iterdir()] == ["Q3 update.pptx"]


def test_bad_name_exits_two(tmp_path: Path) -> None:
    template = build_template(tmp_path / "Template.pptx")
    data = write_data(tmp_path, "CLIENT_NAME: Example\n")
    assert subject.main([str(template), str(data), "--no-pdf", "-q", "-n", "bad/name"]) == 2


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
