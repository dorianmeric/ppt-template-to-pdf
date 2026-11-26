#!/usr/bin/env python3
"""Fill a PowerPoint template from YAML and optionally export the result to PDF.

See README.md for the YAML syntax, the template naming conventions, and the
list of error messages.
"""

from __future__ import annotations

import argparse
import difflib
import re
import sys
import time
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from html.parser import HTMLParser
from importlib.util import find_spec
from pathlib import Path
from typing import Any, Iterable, Iterator, Optional, Sequence

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shapes.group import GroupShape
from pptx.util import Pt

__version__ = "1.1.0"

PP_SAVE_AS_PDF = 32
RETRYABLE_COM_ERRORS = {-2147418111, -2147417846}
EXPORT_ATTEMPTS = 3
MAX_OUTPUT_PATH_LENGTH = 250

DEFAULT_BULLET_CHAR = "•"
DEFAULT_NUMBER_FORMAT = "arabicPeriod"
TRUE_WORDS = {"true", "yes", "on", "1"}
FALSE_WORDS = {"false", "no", "off", "0"}

SUPPORTED_TEMPLATE_SUFFIXES = {".pptx", ".pptm", ".potx"}
SUPPORTED_DATA_SUFFIXES = {".yaml", ".yml"}

HEX_COLOUR_PATTERN = re.compile(r"^#?([0-9A-Fa-f]{6})$")
TOKEN_SCAN_PATTERN = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
FONT_SIZE_PATTERN = re.compile(r"([0-9]+(?:\.[0-9]+)?)(?:pt|px)?", re.IGNORECASE)

ALIGNMENT_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "centre": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}

RUN_STYLE_KEYS = ("bold", "italic", "underline", "colour", "color", "size", "font")
RUN_KEYS = ("text", *RUN_STYLE_KEYS)
PARAGRAPH_KEYS = ("text", "runs", "bullet", "bullet_char", "number", "level", "align", *RUN_STYLE_KEYS)
CONTENT_KEYS = (*PARAGRAPH_KEYS, "paragraphs")

BULLET_ELEMENT_NAMES = (
    "a:buNone",
    "a:buChar",
    "a:buAutoNum",
    "a:buBlip",
    "a:buFont",
    "a:buFontTx",
    "a:buClr",
    "a:buClrTx",
    "a:buSzPct",
    "a:buSzPts",
    "a:buSzTx",
)


class ScriptError(Exception):
    """Base class for failures the script reports without a traceback."""

    exit_code = 1


class InputError(ScriptError):
    """Bad arguments, unreadable files, or invalid YAML content."""

    exit_code = 2


class FieldMatchError(ScriptError):
    """The YAML refers to IDs the template does not contain."""

    exit_code = 3


class ExportError(ScriptError):
    """PowerPoint could not produce the PDF."""

    exit_code = 4


@dataclass
class Reporter:
    """Send progress and diagnostics to the console, honouring -v and -q."""

    quiet: bool = False
    verbose: bool = False

    def info(self, message: str) -> None:
        if not self.quiet:
            print(message)

    def detail(self, message: str) -> None:
        if self.verbose and not self.quiet:
            print(message)

    def warning(self, message: str) -> None:
        if not self.quiet:
            print(f"WARNING: {message}", file=sys.stderr)

    def error(self, message: str) -> None:
        print(f"ERROR: {message}", file=sys.stderr)


@dataclass
class Job:
    """Everything the pipeline needs after the command line has been validated."""

    template: Path
    data: Optional[Path] = None
    output_dir: Optional[Path] = None
    export_pdf: bool = True
    overwrite: bool = False
    dry_run: bool = False
    list_fields: bool = False


@dataclass
class Field:
    """One text shape found in the template."""

    slide: int
    shape: Any
    name: str
    text: str
    matched_by: str = ""


@dataclass
class MatchPlan:
    """The result of matching YAML IDs against template fields."""

    matches: dict[str, list[Field]] = field(default_factory=dict)
    missing: list[str] = field(default_factory=list)

    @property
    def field_count(self) -> int:
        return sum(len(fields) for fields in self.matches.values())


def parse_arguments(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="ppt_template_to_pdf",
        description="Fill a PowerPoint template from YAML and export the result to PDF.",
        epilog=(
            "Examples:\n"
            '  ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"\n'
            '  ppt_template_to_pdf.py "Template.pptx" --list-fields\n'
            '  ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --dry-run\n'
            '  ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" -o Output --no-pdf'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("template", type=Path, help="Path to the template .pptx file.")
    parser.add_argument(
        "yaml_file",
        type=Path,
        nargs="?",
        default=None,
        help="Path to the .yaml or .yml data file. Not needed with --list-fields.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Output folder. Defaults to the template's folder.",
    )
    parser.add_argument("--no-pdf", action="store_true", help="Create only the PPTX file and skip PDF export.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Check the YAML and the template, print the matches, and write nothing.",
    )
    parser.add_argument(
        "--list-fields",
        action="store_true",
        help="List every text shape name and token in the template, then exit.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing output files instead of adding a number to the name.",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Print every field that is filled.")
    parser.add_argument("-q", "--quiet", action="store_true", help="Print errors only.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser.parse_args(argv)


def validate_keys(mapping: dict[str, Any], allowed: Sequence[str], where: str, label: str) -> None:
    """Reject unknown keys so typos surface instead of being ignored."""
    unknown = [key for key in mapping if key not in allowed]
    if not unknown:
        return
    hints = []
    for key in unknown:
        close = difflib.get_close_matches(str(key), list(allowed), n=1)
        hints.append(f"'{key}'" + (f" (did you mean '{close[0]}'?)" if close else ""))
    raise ValueError(
        f"{where}: unknown {label} key(s) {', '.join(hints)}. Allowed keys: {', '.join(allowed)}."
    )


def coerce_bool(value: Any, field_name: str, where: str) -> bool:
    """Accept real booleans plus the string forms people type in YAML."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in TRUE_WORDS:
            return True
        if lowered in FALSE_WORDS:
            return False
    if isinstance(value, int) and value in (0, 1):
        return bool(value)
    raise ValueError(f"{where}: '{field_name}' must be true or false, got {value!r}.")


def coerce_number(
    value: Any,
    field_name: str,
    where: str,
    minimum: float,
    maximum: float,
    integer: bool = False,
) -> float:
    """Convert a YAML value to a bounded number with a readable error."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{where}: '{field_name}' must be a number, got {value!r}.") from None
    if number < minimum or number > maximum:
        raise ValueError(f"{where}: '{field_name}' must be between {minimum:g} and {maximum:g}.")
    if integer:
        if number != int(number):
            raise ValueError(f"{where}: '{field_name}' must be a whole number, got {value!r}.")
        return int(number)
    return number


def normalise_colour(value: Any, where: str = "") -> RGBColor:
    """Convert '#RRGGBB' or 'RRGGBB' text into an RGBColor value."""
    prefix = f"{where}: " if where else ""
    match = HEX_COLOUR_PATTERN.fullmatch(str(value).strip())
    if match is None:
        raise ValueError(f"{prefix}invalid colour {value!r}. Use six hex digits such as '#C00000'.")
    digits = match.group(1)
    return RGBColor(int(digits[0:2], 16), int(digits[2:4], 16), int(digits[4:6], 16))


def parse_font_size(value: Any, where: str = "") -> float:
    """Convert an HTML or CSS size such as '14', '14pt', or '14px' into points."""
    prefix = f"{where}: " if where else ""
    match = FONT_SIZE_PATTERN.fullmatch(str(value).strip())
    if match is None:
        raise ValueError(f"{prefix}invalid font size {value!r}. Use 14 or 14pt.")
    return float(match.group(1))


class PowerPointHTMLParser(HTMLParser):
    """Convert supported HTML tags into PowerPoint paragraph and run definitions."""

    def __init__(self, where: str = "") -> None:
        super().__init__(convert_charrefs=True)
        self.where = where
        self.paragraphs: list[dict[str, Any]] = []
        self.current_paragraph: Optional[dict[str, Any]] = None
        self.styles: list[tuple[str, dict[str, Any]]] = [("__root__", {})]
        self.lists: list[str] = []

    # -- helpers ---------------------------------------------------------

    def _error(self, message: str) -> ValueError:
        prefix = f"{self.where}: " if self.where else ""
        return ValueError(prefix + message)

    def _paragraph(self) -> dict[str, Any]:
        if self.current_paragraph is None:
            self.current_paragraph = {"runs": []}
            self.paragraphs.append(self.current_paragraph)
        return self.current_paragraph

    def _push(self, tag: str, updates: dict[str, Any]) -> None:
        style = dict(self.styles[-1][1])
        style.update(updates)
        self.styles.append((tag, style))

    def _pop(self, tag: str) -> None:
        if len(self.styles) == 1 or self.styles[-1][0] != tag:
            expected = self.styles[-1][0] if len(self.styles) > 1 else "none"
            raise self._error(f"mismatched HTML closing tag </{tag}>; expected </{expected}>.")
        self.styles.pop()

    def _reject_attributes(self, tag: str, attributes: dict[str, Any], allowed: set[str]) -> None:
        unsupported = set(attributes) - allowed
        if unsupported:
            raise self._error(f"unsupported <{tag}> attributes: {sorted(unsupported)}")

    # -- start tags ------------------------------------------------------

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        attributes = {name.lower(): value for name, value in attrs}
        handler = getattr(self, f"_start_{tag.lower()}", None)
        if handler is None:
            raise self._error(
                f"unsupported HTML tag <{tag}>. Supported tags: "
                "b, strong, u, font, span, ul, ol, li, br, p. "
                "Use &lt; and &gt; to print a tag literally."
            )
        handler(attributes)

    def _start_b(self, attributes: dict[str, Any]) -> None:
        self._push("b", {"bold": True})

    def _start_strong(self, attributes: dict[str, Any]) -> None:
        self._push("strong", {"bold": True})

    def _start_u(self, attributes: dict[str, Any]) -> None:
        self._push("u", {"underline": True})

    def _start_font(self, attributes: dict[str, Any]) -> None:
        self._reject_attributes("font", attributes, {"color", "colour", "size"})
        updates: dict[str, Any] = {}
        colour = attributes.get("color", attributes.get("colour"))
        if colour is not None:
            normalise_colour(colour, self.where)
            updates["colour"] = colour
        if attributes.get("size") is not None:
            updates["size"] = parse_font_size(attributes["size"], self.where)
        self._push("font", updates)

    def _start_span(self, attributes: dict[str, Any]) -> None:
        self._reject_attributes("span", attributes, {"style"})
        self._push("span", self._css(attributes.get("style") or ""))

    def _css(self, value: str) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for declaration in value.split(";"):
            if not declaration.strip():
                continue
            if ":" not in declaration:
                raise self._error(f"invalid CSS declaration {declaration.strip()!r}.")
            name, css_value = (part.strip() for part in declaration.split(":", 1))
            name = name.lower()
            if name in {"color", "colour"}:
                normalise_colour(css_value, self.where)
                result["colour"] = css_value
            elif name == "font-size":
                result["size"] = parse_font_size(css_value, self.where)
            else:
                raise self._error(
                    f"unsupported CSS property '{name}'; use color or font-size. "
                    "For other formatting use the YAML run fields instead."
                )
        return result

    def _start_ul(self, attributes: dict[str, Any]) -> None:
        self.lists.append("ul")

    def _start_ol(self, attributes: dict[str, Any]) -> None:
        self.lists.append("ol")

    def _start_li(self, attributes: dict[str, Any]) -> None:
        if not self.lists:
            raise self._error("HTML <li> must be inside <ul> or <ol>.")
        definition: dict[str, Any] = {"runs": [], "level": len(self.lists) - 1}
        if self.lists[-1] == "ol":
            definition["number"] = True
        else:
            definition["bullet"] = True
        self.current_paragraph = definition
        self.paragraphs.append(definition)

    def _start_br(self, attributes: dict[str, Any]) -> None:
        self.current_paragraph = None

    def _start_p(self, attributes: dict[str, Any]) -> None:
        self.current_paragraph = None

    # -- end tags --------------------------------------------------------

    def handle_endtag(self, tag: str) -> None:
        handler = getattr(self, f"_end_{tag.lower()}", None)
        if handler is None:
            raise self._error(f"unsupported HTML closing tag </{tag}>.")
        handler()

    def _end_b(self) -> None:
        self._pop("b")

    def _end_strong(self) -> None:
        self._pop("strong")

    def _end_u(self) -> None:
        self._pop("u")

    def _end_font(self) -> None:
        self._pop("font")

    def _end_span(self) -> None:
        self._pop("span")

    def _end_li(self) -> None:
        self.current_paragraph = None

    def _end_ul(self) -> None:
        self._close_list("ul")

    def _end_ol(self) -> None:
        self._close_list("ol")

    def _end_br(self) -> None:
        self.current_paragraph = None

    def _end_p(self) -> None:
        self.current_paragraph = None

    def _close_list(self, tag: str) -> None:
        if not self.lists or self.lists[-1] != tag:
            raise self._error(f"mismatched HTML closing tag </{tag}>.")
        self.lists.pop()
        self.current_paragraph = None

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        if tag.lower() != "br":
            raise self._error(f"unsupported self-closing HTML tag <{tag}/>.")
        self.current_paragraph = None

    def handle_data(self, data: str) -> None:
        if data == "":
            return
        if data.strip() == "" and ("\n" in data or "\r" in data):
            return
        run: dict[str, Any] = {"text": data}
        run.update(self.styles[-1][1])
        self._paragraph()["runs"].append(run)

    def validate_complete(self) -> None:
        if len(self.styles) != 1:
            raise self._error(f"unclosed HTML formatting tag <{self.styles[-1][0]}>.")
        if self.lists:
            raise self._error(f"unclosed HTML list tag <{self.lists[-1]}>.")


def parse_html_content(value: str, where: str = "") -> list[dict[str, Any]]:
    """Parse real HTML tags and keep entity-escaped tags as literal text."""
    parser = PowerPointHTMLParser(where)
    parser.feed(value)
    parser.close()
    parser.validate_complete()
    return parser.paragraphs or [{"runs": [{"text": ""}]}]


def value_to_paragraph_definitions(value: Any, where: str = "") -> list[dict[str, Any]]:
    """Normalise any YAML value into a list of paragraph definitions."""
    if value is None:
        return [{"runs": [{"text": ""}]}]
    if isinstance(value, str):
        return parse_html_content(value, where)
    if isinstance(value, bool):
        return [{"runs": [{"text": "true" if value else "false"}]}]
    if isinstance(value, (int, float)):
        return [{"runs": [{"text": str(value)}]}]
    if not isinstance(value, dict):
        raise ValueError(f"{where}: unsupported content {value!r}. Use text, runs, or paragraphs.")
    validate_keys(value, CONTENT_KEYS, where, "content")
    if "paragraphs" in value:
        if "runs" in value or "text" in value:
            raise ValueError(f"{where}: use either 'paragraphs' or 'text'/'runs', not both.")
        paragraphs = value["paragraphs"]
        if not isinstance(paragraphs, list):
            raise ValueError(f"{where}: 'paragraphs' must be a list.")
        return paragraphs
    if "runs" in value:
        return [dict(value)]
    if "text" in value:
        paragraph = dict(value)
        run = {"text": paragraph.pop("text")}
        for key in RUN_STYLE_KEYS:
            if key in paragraph:
                run[key] = paragraph.pop(key)
        paragraph["runs"] = [run]
        return [paragraph]
    raise ValueError(f"{where}: content mapping needs 'text', 'runs', or 'paragraphs'.")


def remove_bullet_elements(paragraph) -> None:
    """Remove inherited bullet and numbering elements from a paragraph."""
    properties = paragraph._p.get_or_add_pPr()
    for element_name in BULLET_ELEMENT_NAMES:
        for element in properties.findall(qn(element_name)):
            properties.remove(element)


def set_paragraph_marker(paragraph, marker: Optional[str], character: str = DEFAULT_BULLET_CHAR) -> None:
    """Set a paragraph to a bullet, a number, or no marker at all."""
    remove_bullet_elements(paragraph)
    properties = paragraph._p.get_or_add_pPr()
    if marker == "bullet":
        element = OxmlElement("a:buChar")
        element.set("char", character)
    elif marker == "number":
        element = OxmlElement("a:buAutoNum")
        element.set("type", DEFAULT_NUMBER_FORMAT)
    else:
        element = OxmlElement("a:buNone")
    properties.insert_element_before(element, "a:tabLst", "a:defRPr", "a:extLst")


def copy_paragraph_defaults(source_paragraph, target_paragraph) -> None:
    """Copy paragraph-level properties from a template paragraph."""
    source_properties = source_paragraph._p.pPr
    if source_properties is None:
        return
    target_properties = target_paragraph._p.pPr
    if target_properties is not None:
        target_paragraph._p.remove(target_properties)
    target_paragraph._p.insert(0, deepcopy(source_properties))


def copy_run_defaults(source_run, target_run) -> None:
    """Copy character formatting from a template run."""
    source_properties = source_run._r.rPr
    if source_properties is None:
        return
    target_properties = target_run._r.rPr
    if target_properties is not None:
        target_run._r.remove(target_properties)
    target_run._r.insert(0, deepcopy(source_properties))


def apply_run_format(run, run_definition: dict[str, Any], where: str) -> None:
    """Apply bold, italic, underline, colour, size, and font to one run."""
    if "bold" in run_definition:
        run.font.bold = coerce_bool(run_definition["bold"], "bold", where)
    if "italic" in run_definition:
        run.font.italic = coerce_bool(run_definition["italic"], "italic", where)
    if "underline" in run_definition:
        run.font.underline = coerce_bool(run_definition["underline"], "underline", where)
    if "colour" in run_definition or "color" in run_definition:
        colour_value = run_definition.get("colour", run_definition.get("color"))
        run.font.color.rgb = normalise_colour(colour_value, where)
    if "size" in run_definition:
        run.font.size = Pt(coerce_number(run_definition["size"], "size", where, 1, 400))
    if "font" in run_definition:
        font_name = str(run_definition["font"]).strip()
        if not font_name:
            raise ValueError(f"{where}: 'font' must not be empty.")
        run.font.name = font_name


def apply_paragraph_format(paragraph, definition: dict[str, Any], where: str) -> None:
    """Apply level, alignment, and bullet or numbering settings to a paragraph."""
    paragraph.level = coerce_number(definition.get("level", 0), "level", where, 0, 8, integer=True)
    if "align" in definition:
        alignment_name = str(definition["align"]).strip().lower()
        if alignment_name not in ALIGNMENT_MAP:
            raise ValueError(
                f"{where}: unsupported alignment {definition['align']!r}; "
                f"use {', '.join(sorted(ALIGNMENT_MAP))}."
            )
        paragraph.alignment = ALIGNMENT_MAP[alignment_name]
    bullet = definition.get("bullet")
    number = definition.get("number")
    if bullet is not None and number is not None:
        raise ValueError(f"{where}: use 'bullet' or 'number', not both.")
    if bullet is not None:
        character = str(definition.get("bullet_char", DEFAULT_BULLET_CHAR))
        if len(character) != 1:
            raise ValueError(f"{where}: 'bullet_char' must be a single character, got {character!r}.")
        enabled = coerce_bool(bullet, "bullet", where)
        set_paragraph_marker(paragraph, "bullet" if enabled else None, character)
    elif number is not None:
        enabled = coerce_bool(number, "number", where)
        set_paragraph_marker(paragraph, "number" if enabled else None)
    elif "bullet_char" in definition:
        raise ValueError(f"{where}: 'bullet_char' needs 'bullet: true'.")


def populate_text_frame(text_frame, content_definition: Any, where: str = "text frame") -> None:
    """Replace a text frame with YAML-defined paragraphs, bullets, and runs."""
    source_paragraph = text_frame.paragraphs[0] if text_frame.paragraphs else None
    source_run = source_paragraph.runs[0] if source_paragraph is not None and source_paragraph.runs else None
    paragraph_definitions = value_to_paragraph_definitions(content_definition, where)
    text_frame.clear()
    for paragraph_index, paragraph_definition in enumerate(paragraph_definitions):
        paragraph_where = f"{where}, paragraph {paragraph_index + 1}"
        if not isinstance(paragraph_definition, dict):
            raise ValueError(f"{paragraph_where}: must be a mapping, got {paragraph_definition!r}.")
        validate_keys(paragraph_definition, PARAGRAPH_KEYS, paragraph_where, "paragraph")
        paragraph = text_frame.paragraphs[0] if paragraph_index == 0 else text_frame.add_paragraph()
        if source_paragraph is not None:
            copy_paragraph_defaults(source_paragraph, paragraph)
        apply_paragraph_format(paragraph, paragraph_definition, paragraph_where)
        run_definitions = paragraph_definition.get("runs", [{"text": ""}])
        if not isinstance(run_definitions, list):
            raise ValueError(f"{paragraph_where}: 'runs' must be a list.")
        for run_index, run_definition in enumerate(run_definitions):
            run_where = f"{paragraph_where}, run {run_index + 1}"
            if isinstance(run_definition, (str, int, float, bool)):
                run_definition = {"text": str(run_definition)}
            if not isinstance(run_definition, dict):
                raise ValueError(f"{run_where}: must be text or a mapping, got {run_definition!r}.")
            validate_keys(run_definition, RUN_KEYS, run_where, "run")
            raw_text = run_definition.get("text", "")
            if raw_text is not None and not isinstance(raw_text, (str, int, float, bool)):
                raise ValueError(f"{run_where}: 'text' must be text, got {raw_text!r}.")
            run = paragraph.add_run()
            if source_run is not None:
                copy_run_defaults(source_run, run)
            run.text = "" if raw_text is None else str(raw_text)
            apply_run_format(run, run_definition, run_where)


def iter_shapes(shapes: Iterable[Any]) -> Iterator[Any]:
    """Yield every shape, including shapes inside groups."""
    for shape in shapes:
        yield shape
        if isinstance(shape, GroupShape):
            yield from iter_shapes(shape.shapes)


def collect_fields(presentation: Presentation) -> list[Field]:
    """Collect every text shape in the presentation with its slide number."""
    fields: list[Field] = []
    for slide_number, slide in enumerate(presentation.slides, start=1):
        for shape in iter_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            fields.append(
                Field(
                    slide=slide_number,
                    shape=shape,
                    name=(getattr(shape, "name", "") or "").strip(),
                    text=getattr(shape, "text", "") or "",
                )
            )
    return fields


def token_pattern(unique_id: str) -> re.Pattern[str]:
    """Build a tolerant matcher for '{{ID}}' with optional inner spaces."""
    return re.compile(r"\{\{\s*" + re.escape(unique_id.strip()) + r"\s*\}\}", re.IGNORECASE)


def match_reason(field: Field, unique_id: str, pattern: re.Pattern[str]) -> Optional[str]:
    """Return how a field matches an ID, or None when it does not match."""
    key = unique_id.strip()
    if field.name == key:
        return "shape name"
    if field.name.casefold() == key.casefold():
        return "shape name (case-insensitive)"
    if pattern.search(field.text):
        return "token"
    return None


def resolve_plan(fields: list[Field], yaml_data: dict[str, Any], reporter: Reporter) -> MatchPlan:
    """Match every YAML ID against the template fields before anything is written."""
    plan = MatchPlan()
    owners: dict[int, str] = {}
    for unique_id in yaml_data:
        pattern = token_pattern(unique_id)
        matched = []
        for template_field in fields:
            reason = match_reason(template_field, unique_id, pattern)
            if reason is None:
                continue
            template_field.matched_by = reason
            if "insensitive" in reason:
                reporter.warning(
                    f"'{unique_id}' matched slide {template_field.slide} "
                    f"shape '{template_field.name}' by {reason}."
                )
            previous = owners.get(id(template_field.shape))
            if previous is not None and previous != unique_id:
                reporter.warning(
                    f"slide {template_field.slide} shape '{template_field.name}' is targeted by "
                    f"'{previous}' and '{unique_id}'; '{unique_id}' wins."
                )
            owners[id(template_field.shape)] = unique_id
            matched.append(template_field)
        if matched:
            plan.matches[unique_id] = matched
        else:
            plan.missing.append(unique_id)
    return plan


def build_missing_message(missing: Sequence[str], fields: Sequence[Field]) -> str:
    """Explain which IDs were not found and suggest near matches."""
    candidates = {template_field.name for template_field in fields if template_field.name}
    for template_field in fields:
        candidates.update(match.strip() for match in TOKEN_SCAN_PATTERN.findall(template_field.text))
    lines = ["These YAML IDs were not found in the template:"]
    for unique_id in missing:
        close = difflib.get_close_matches(unique_id, sorted(candidates), n=1, cutoff=0.7)
        suffix = f" (did you mean '{close[0]}'?)" if close else ""
        lines.append(f"  - {unique_id}{suffix}")
    lines.append("")
    lines.append("Run the script with --list-fields to see every name and token in the template.")
    return "\n".join(lines)


def fill_presentation(plan: MatchPlan, yaml_data: dict[str, Any], reporter: Reporter) -> None:
    """Write the YAML content into every matched field."""
    for unique_id, matched in plan.matches.items():
        for template_field in matched:
            reporter.detail(
                f"  {unique_id} -> slide {template_field.slide} "
                f"'{template_field.name}' ({template_field.matched_by})"
            )
            populate_text_frame(
                template_field.shape.text_frame,
                yaml_data[unique_id],
                where=f"YAML '{unique_id}' (slide {template_field.slide}, shape '{template_field.name}')",
            )


def print_field_list(fields: Sequence[Field]) -> None:
    """Print every text shape so the user can build the YAML file."""
    if not fields:
        print("No text shapes found in this presentation.")
        return
    rows = []
    for template_field in fields:
        token = TOKEN_SCAN_PATTERN.search(template_field.text)
        source = f"token {token.group(1).strip()}" if token else "shape name"
        text = " ".join(template_field.text.split())
        if len(text) > 60:
            text = text[:57] + "..."
        rows.append((template_field, source, text))
    name_width = max(len("Shape name"), *(len(template_field.name) for template_field, _, _ in rows))
    source_width = max(len("Match by"), *(len(source) for _, source, _ in rows))
    print(f"{'Slide':<7}{'Shape name':<{name_width}}  {'Match by':<{source_width}}  Text")
    print(f"{'-' * 7}{'-' * name_width}  {'-' * source_width}  {'-' * 20}")
    for template_field, source, text in rows:
        print(f"{template_field.slide:<7}{template_field.name:<{name_width}}  {source:<{source_width}}  {text}")
    print()
    print("Use these shape names, or the IDs inside {{braces}}, as top-level YAML keys.")


def load_yaml_data(yaml_path: Path) -> dict[str, Any]:
    """Load YAML safely and report syntax errors with line, column, and a fix."""
    try:
        yaml_text = yaml_path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as error:
        raise InputError(
            f"{yaml_path} is not valid UTF-8 text ({error.reason}). Save the file as UTF-8 and try again."
        ) from None
    lint_yaml_text(yaml_text, yaml_path)
    try:
        loaded_data = yaml.load(yaml_text, Loader=UniqueKeyLoader)
    except yaml.YAMLError as error:
        raise InputError(format_yaml_error(yaml_path, yaml_text, error)) from None
    if loaded_data is None:
        raise InputError(f"{yaml_path} is empty. Add at least one 'ID: content' entry.")
    if not isinstance(loaded_data, dict):
        raise InputError(
            f"{yaml_path}: the top level must be a mapping of 'ID: content', "
            f"not {type(loaded_data).__name__}."
        )
    data: dict[str, Any] = {}
    for key, content in loaded_data.items():
        if not isinstance(key, str):
            raise InputError(
                f"{yaml_path}: ID {key!r} is not text. YAML reads words such as NO, ON, YES, "
                "and dates as true/false or dates. Wrap the ID in quotes, for example 'NO': content."
            )
        data[key.strip()] = content
    if not data:
        raise InputError(f"{yaml_path} contains no IDs.")
    return data


class UniqueKeyLoader(yaml.SafeLoader):
    """SafeLoader that refuses duplicate mapping keys instead of keeping the last one."""

    def construct_mapping(self, node, deep: bool = False) -> dict[Any, Any]:
        seen: set[Any] = set()
        for key_node, _value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            try:
                duplicate = key in seen
            except TypeError:
                continue
            if duplicate:
                raise yaml.constructor.ConstructorError(
                    "while constructing a mapping",
                    node.start_mark,
                    f"found duplicate key {key!r}",
                    key_node.start_mark,
                )
            seen.add(key)
        return super().construct_mapping(node, deep=deep)


def format_yaml_excerpt(lines: Sequence[str], line_index: int, column_index: int) -> str:
    """Quote the failing line with neighbours and a caret under the column."""
    first = max(0, line_index - 2)
    last = min(len(lines), line_index + 2)
    width = len(str(last))
    excerpt: list[str] = []
    for index in range(first, last):
        prefix = ">" if index == line_index else " "
        excerpt.append(f"{prefix} {index + 1:>{width}} | {lines[index]}")
        if index == line_index:
            excerpt.append(" " * (width + 5 + column_index) + "^")
    return "\n".join(excerpt)


def single_quoted_rewrite(source_line: str) -> Optional[str]:
    """Rewrite 'KEY: "value"' as 'KEY: 'value'' so inner double quotes survive."""
    double_quote = '"'
    single_quote = "'"
    if ":" not in source_line:
        return None
    key, raw_value = source_line.split(":", 1)
    value = raw_value.strip()
    if not (value.startswith(double_quote) and value.endswith(double_quote)):
        return None
    value = value[1:-1].replace(single_quote, single_quote * 2)
    return f"{key}: '{value}'"


def yaml_error_suggestion(source_line: str, problem: str) -> str:
    """Turn a PyYAML problem description into a concrete correction."""
    if "duplicate key" in problem.lower():
        return "This key appears twice in the same block. Remove the duplicate or rename one of them."
    indentation = source_line[: len(source_line) - len(source_line.lstrip())]
    if "\t" in indentation:
        return "The indentation contains a TAB. YAML needs spaces; replace the tab with spaces."
    rewrite = single_quoted_rewrite(source_line)
    if rewrite is not None and "<" in source_line and ">" in source_line:
        return (
            "This value is wrapped in double quotes but contains double-quoted HTML attributes. "
            'For example size="14" ends the YAML string early.\n\n'
            "Recommended change (copy/paste):\n"
            f"    {rewrite}\n\n"
            "For longer HTML content, a YAML block value using '|' also works."
        )
    if source_line.lstrip().startswith("-"):
        return "This is a YAML list item. Align '-' with its siblings and indent its child fields further."
    return "Check the highlighted line and the line above it for mismatched quotes, indentation, or a missing colon."


def find_unquoted_hash(line: str) -> Optional[int]:
    """Return the index of a '#' that YAML would treat as a comment."""
    quote = ""
    index = 0
    while index < len(line):
        character = line[index]
        if quote:
            if character == "\\" and quote == '"':
                index += 2
                continue
            if character == quote:
                quote = ""
        elif character in "\"'":
            quote = character
        elif character == "#":
            return index
        index += 1
    return None


def lint_yaml_text(yaml_text: str, yaml_path: Path) -> None:
    """Catch HTML values that YAML silently truncates at a '#' comment."""
    for line_number, line in enumerate(yaml_text.splitlines(), start=1):
        hash_index = find_unquoted_hash(line)
        if hash_index is None or hash_index == 0:
            continue
        before, after = line[:hash_index], line[hash_index:]
        if not re.search(r"=\s*['\"]", before) or '"' not in after:
            continue
        rewrite = single_quoted_rewrite(line)
        fix = (
            f"    {rewrite}\n"
            if rewrite is not None
            else "    Wrap the value in single quotes, or use a '|' block value.\n"
        )
        raise InputError(
            "YAML COMMENT ERROR\n"
            "==================\n"
            f"File:    {yaml_path}\n"
            f"Line:    {line_number}\n"
            f"Problem: YAML reads '#' as the start of a comment, so this value stops at "
            f"{after.split('\"')[0].split()[0]!r}.\n\n"
            f"> {line_number:>3} | {line}\n\n"
            "How to fix it:\n"
            "Wrap the whole value in single quotes so the inner double quotes and the '#' survive.\n"
            f"{fix}"
        )


def format_yaml_error(yaml_path: Path, yaml_text: str, error: yaml.YAMLError) -> str:
    """Format a PyYAML error with location, source context, and a likely correction."""
    marker = getattr(error, "problem_mark", None) or getattr(error, "context_mark", None)
    problem = getattr(error, "problem", None) or str(error).splitlines()[0]
    if marker is None:
        return (
            "YAML SYNTAX ERROR\n"
            "=================\n"
            f"File:    {yaml_path}\n"
            f"Problem: {problem}\n\n"
            "Check quotes, indentation, colons, and HTML attributes."
        )
    lines = yaml_text.splitlines()
    source_line = lines[marker.line] if 0 <= marker.line < len(lines) else ""
    return (
        "YAML SYNTAX ERROR\n"
        "=================\n"
        f"File:    {yaml_path}\n"
        f"Line:    {marker.line + 1}\n"
        f"Column:  {marker.column + 1}\n"
        f"Problem: {problem}\n\n"
        f"{format_yaml_excerpt(lines, marker.line, marker.column)}\n\n"
        "How to fix it:\n"
        f"{yaml_error_suggestion(source_line, problem)}"
    )


def resolve_template(path: Path) -> Path:
    """Validate the template path."""
    path = path.expanduser().resolve()
    if not path.is_file():
        raise InputError(f"Template file not found: {path}")
    if path.suffix.lower() not in SUPPORTED_TEMPLATE_SUFFIXES:
        raise InputError(
            f"Template must be one of {', '.join(sorted(SUPPORTED_TEMPLATE_SUFFIXES))}: {path.name}"
        )
    return path


def resolve_data_file(path: Path) -> Path:
    """Validate the YAML path."""
    path = path.expanduser().resolve()
    if not path.is_file():
        raise InputError(f"YAML file not found: {path}")
    if path.suffix.lower() not in SUPPORTED_DATA_SUFFIXES:
        raise InputError(f"Data file must be .yaml or .yml: {path.name}")
    return path


def resolve_output_dir(path: Optional[Path], template: Path, create: bool) -> Path:
    """Validate or create the output folder."""
    output_dir = (path or template.parent).expanduser().resolve()
    if output_dir.exists() and not output_dir.is_dir():
        raise InputError(f"Output path is a file, not a folder: {output_dir}")
    if create:
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            raise InputError(
                f"Cannot create output folder {output_dir}: {error.strerror or error}"
            ) from None
    return output_dir


def resolve_output_paths(job: Job, reporter: Reporter) -> tuple[Path, Path]:
    """Build output paths, never clobbering earlier files unless --overwrite is set."""
    assert job.data is not None and job.output_dir is not None
    stem = f"{job.template.stem} - {job.data.stem} - {date.today().isoformat()}"
    targets = [job.output_dir / f"{stem}.pptx"]
    if job.export_pdf:
        targets.append(job.output_dir / f"{stem}.pdf")
    if job.overwrite:
        return targets[0], targets[-1]
    for index in range(1, 1000):
        suffix = "" if index == 1 else f" ({index})"
        pptx_path = job.output_dir / f"{stem}{suffix}.pptx"
        pdf_path = job.output_dir / f"{stem}{suffix}.pdf"
        candidate = [pptx_path, pdf_path] if job.export_pdf else [pptx_path]
        if not any(path.exists() for path in candidate):
            if index > 1:
                reporter.warning(
                    f"{targets[0].name} already exists. Writing {pptx_path.name} instead. "
                    "Use --overwrite to replace the earlier file."
                )
            return pptx_path, pdf_path
    raise InputError(f"Cannot find a free output name for {stem} in {job.output_dir}")


def check_path_length(paths: Iterable[Path], reporter: Reporter) -> None:
    """Warn before Windows rejects a long output path."""
    for path in paths:
        if len(str(path)) > MAX_OUTPUT_PATH_LENGTH:
            reporter.warning(
                f"Output path is {len(str(path))} characters long and Windows may refuse it: {path}. "
                "Use a shorter --output-dir."
            )


def open_presentation(path: Path) -> Presentation:
    """Open a template and explain the common failures."""
    try:
        return Presentation(str(path))
    except Exception as error:
        raise InputError(
            f"Cannot open template {path}: {error}. "
            "Check that the file is a valid PowerPoint file and is not password protected."
        ) from error


def save_presentation(presentation: Presentation, output_path: Path) -> None:
    """Save the filled presentation and explain the common failures."""
    try:
        presentation.save(str(output_path))
    except PermissionError as error:
        raise InputError(
            f"Cannot write {output_path}. Close the file in PowerPoint and check folder permissions."
        ) from error
    except OSError as error:
        raise InputError(f"Cannot write {output_path}: {error.strerror or error}") from error


def com_hresult(error: BaseException) -> Optional[int]:
    """Extract the HRESULT from a pywin32 COM error."""
    hresult = getattr(error, "hresult", None)
    if hresult is None and error.args and isinstance(error.args[0], int):
        hresult = error.args[0]
    return hresult


def is_retryable_com_error(error: BaseException) -> bool:
    """PowerPoint rejects calls while it is busy; those errors are worth retrying."""
    return com_hresult(error) in RETRYABLE_COM_ERRORS


def describe_export_error(error: BaseException, pptx_path: Path) -> str:
    """Translate a COM failure into an action the user can take."""
    message = str(error)
    lowered = message.lower()
    if any(text in lowered for text in ("class string", "class not registered", "invalid class")):
        hint = (
            "PowerPoint automation is not registered. Install desktop PowerPoint, then run "
            'python "<python>\\Scripts\\pywin32_postinstall.py" -install'
        )
    elif "could not open" in lowered or "cannot open" in lowered:
        hint = (
            f"PowerPoint could not open {pptx_path}. Close the file, move it out of OneDrive or a "
            "network share, and unblock it in the file Properties."
        )
    elif isinstance(error, PermissionError) or "permission" in lowered:
        hint = "The output PDF is locked. Close it in a PDF reader and try again."
    else:
        hint = "Check that PowerPoint opens normally, then run the script again."
    return f"PDF export failed: {message}\n{hint}"


def save_as_pdf(win32_client: Any, pptx_path: Path, pdf_path: Path) -> None:
    """Open the PPTX in a hidden PowerPoint instance and save it as PDF."""
    application = None
    presentation = None
    try:
        application = win32_client.DispatchEx("PowerPoint.Application")
        presentation = application.Presentations.Open(
            str(pptx_path.resolve()),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        presentation.SaveAs(str(pdf_path.resolve()), PP_SAVE_AS_PDF)
    finally:
        for closer in (
            lambda: presentation.Close() if presentation is not None else None,
            lambda: application.Quit() if application is not None else None,
        ):
            try:
                closer()
            except Exception:
                pass


def export_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> None:
    """Export a PPTX to PDF with the installed Windows PowerPoint application."""
    if sys.platform != "win32":
        raise ExportError("PDF export needs Windows and desktop PowerPoint. Add --no-pdf to skip it.")
    try:
        import win32com.client
    except ImportError as error:
        raise ExportError(
            "pywin32 is required for PDF export. Install it with: python -m pip install pywin32. "
            "Add --no-pdf to skip PDF export."
        ) from error
    for attempt in range(1, EXPORT_ATTEMPTS + 1):
        try:
            save_as_pdf(win32com.client, pptx_path, pdf_path)
            return
        except Exception as error:
            if attempt == EXPORT_ATTEMPTS or not is_retryable_com_error(error):
                raise ExportError(describe_export_error(error, pptx_path)) from error
            time.sleep(1.5 * attempt)
    raise ExportError(f"PDF export failed after {EXPORT_ATTEMPTS} attempts.")


def build_job(arguments: argparse.Namespace) -> Job:
    """Validate the command line and turn it into a job."""
    template = resolve_template(arguments.template)
    if arguments.list_fields:
        return Job(template=template, output_dir=template.parent, list_fields=True)
    if arguments.yaml_file is None:
        raise InputError(
            "A YAML data file is required. Pass it as the second argument, "
            "or use --list-fields to inspect the template first."
        )
    data = resolve_data_file(arguments.yaml_file)
    export_pdf = not arguments.no_pdf
    if export_pdf:
        if sys.platform != "win32":
            raise InputError("PDF export needs Windows. Add --no-pdf to create only the PPTX file.")
        if find_spec("win32com") is None:
            raise InputError(
                "PDF export needs pywin32. Install it with: python -m pip install pywin32. "
                "Add --no-pdf to create only the PPTX file."
            )
    output_dir = resolve_output_dir(arguments.output_dir, template, create=not arguments.dry_run)
    return Job(
        template=template,
        data=data,
        output_dir=output_dir,
        export_pdf=export_pdf,
        overwrite=arguments.overwrite,
        dry_run=arguments.dry_run,
    )


def run_job(job: Job, reporter: Reporter) -> int:
    """Run the whole pipeline for a validated job."""
    if job.list_fields:
        print_field_list(collect_fields(open_presentation(job.template)))
        return 0
    assert job.data is not None
    yaml_data = load_yaml_data(job.data)
    presentation = open_presentation(job.template)
    fields = collect_fields(presentation)
    plan = resolve_plan(fields, yaml_data, reporter)
    if plan.missing:
        raise FieldMatchError(build_missing_message(plan.missing, fields))
    unfilled = [template_field for template_field in fields if not template_field.matched_by]
    for template_field in unfilled:
        reporter.detail(f"  left unchanged: slide {template_field.slide} '{template_field.name}'")
    reporter.detail(f"  {len(plan.matches)} YAML ID(s) matched {plan.field_count} template field(s)")
    if job.dry_run:
        reporter.info(
            f"Dry run: {len(plan.matches)} YAML ID(s) matched {plan.field_count} template field(s). "
            "Nothing was written."
        )
        return 0
    output_pptx, output_pdf = resolve_output_paths(job, reporter)
    check_path_length([output_pptx, output_pdf] if job.export_pdf else [output_pptx], reporter)
    try:
        fill_presentation(plan, yaml_data, reporter)
    except ValueError as error:
        raise InputError(str(error)) from None
    save_presentation(presentation, output_pptx)
    reporter.info(f"Created PPTX: {output_pptx}")
    if job.export_pdf:
        export_pptx_to_pdf(output_pptx, output_pdf)
        reporter.info(f"Created PDF:  {output_pdf}")
    else:
        reporter.detail("  PDF export skipped (--no-pdf)")
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Parse arguments, run the job, and map failures onto exit codes."""
    arguments = parse_arguments(argv)
    reporter = Reporter(quiet=arguments.quiet, verbose=arguments.verbose)
    try:
        return run_job(build_job(arguments), reporter)
    except ScriptError as error:
        reporter.error(str(error))
        return error.exit_code
    except KeyboardInterrupt:
        reporter.error("Interrupted.")
        return 130
    except Exception as error:  # noqa: BLE001 - last resort, keeps the CLI readable
        reporter.error(f"Unexpected failure: {error}")
        return 1


def cli() -> int:
    """Console-script entry point."""
    return main()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    sys.exit(main())
