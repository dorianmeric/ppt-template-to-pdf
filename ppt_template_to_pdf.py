#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Populate a PowerPoint template from YAML, save the populated PPTX, and export it to PDF.

Template convention:
1. Preferred: in PowerPoint's Selection Pane, rename each target text box/placeholder to its unique ID.
2. Alternative: put a token such as {{CLIENT_NAME}} inside a text box.

Required packages:
    python -m pip install python-pptx PyYAML pywin32

Manual package pages / downloads:
    https://pypi.org/project/python-pptx/
    https://pypi.org/project/PyYAML/
    https://pypi.org/project/pywin32/
    
    resulting packages (including dependencies)
    Package           Version
    ----------------- -------
    lxml              6.1.3
    pillow            13.0.0
    pip               26.3
    python-pptx       1.0.2
    pywin32           312
    PyYAML            6.0.4
    typing_extensions 4.16.0
    xlsxwriter        3.3.4
    
Runtime requirements:
- Windows.
- A locally installed desktop version of Microsoft PowerPoint.
- Python 3.10 or later is recommended.

Example command:
    python pptx_template_to_pdf.py.py "Template.pptx" "Deal data.yaml" --output-dir "Output"




# Preparing the template
    The recommended approach is:
    - Open the template in PowerPoint.
    - Open Home → Arrange → Selection Pane.
    - Rename each text box or placeholder using its unique ID, for example:
        CLIENT_NAME
        TRANSACTION_TITLE
        KEY_POINTS
    - Use those exact IDs as the top-level YAML keys.

    Alternatively, place a token such as {{CLIENT_NAME}} inside a text box. The script supports both methods.

    Example YAML
    ```

    CLIENT_NAME: "Example Energy Pte. Ltd."

    TRANSACTION_TITLE:
      runs:
        - text: "Project "
        - text: "Sophia"
          bold: true
          colour: "#C00000"
        - text: " Financing Proposal"

    SUMMARY:
      paragraphs:
        - runs:
            - text: "THEBANK "
              bold: true
            - text: "is proposing a structured financing solution."
        - bullet: true
          level: 0
          runs:
            - text: "Senior facility: "
              bold: true
            - text: "USD 150 million"
              colour: "#C00000"
        - bullet: true
          level: 0
          runs:
            - text: "Expected tenor: "
              bold: true
            - text: "10 years"
              underline: true
        - bullet: true
          level: 1
          runs:
            - text: "Subject to credit approval"
              italic: true
              colour: "#666666"

    DISCLAIMER:
      text: "Strictly private and confidential"
      bold: true
      underline: true
      colour: "#C00000"
      size: 9
      align: "center"

    TITLE: "<b>Bold text</b> and <u>underlined text</u>"

    COLOURED_TEXT: '<font color="#C00000">Red text</font>'

    SIZED_TEXT: '<font size="14">14-point text</font>'

    COMBINED: '<span style="color:#C00000; font-size:14pt"><b>Red, bold, 14-point text</b></span>'

    BULLETS: |
      <ul>
        <li>First bullet</li>
        <li>Second bullet with <b>bold text</b></li>
        <li>Third bullet with <u>underlined text</u></li>
      </ul>

    NESTED_BULLETS: |
      <ul>
        <li>Top-level bullet</li>
        <li>Another top-level bullet
          <ul>
            <li>Nested bullet</li>
          </ul>
        </li>
      </ul>


    (to print HTML text in the final ppt)
    LITERAL_TAG: "&lt;b&gt;This is not bold&lt;/b&gt;"
    LITERAL_TAG: "&#60;u&#62;This is not underlined&#60;/u&#62;"

    ``` 
  
  
Supported run formatting fields are:
	bold
	italic
	underline
	colour or color, using#RRGGBB
	size, in points
	font

Supported paragraph fields include:
	bullet
	bullet_char
	level
	align
	runs


Commands:
python pptx_template_to_pdf.py.py "C:\Input\Template.pptx" "C:\Input\Deal data.yaml"
python pptx_template_to_pdf.py.py "C:\Input\Template.pptx" "C:\Input\Deal data.yaml" --output-dir "C:\Output"

Output names for a run on 2026-11-26:
    Template - Deal data - 2026-11-26.pptx
    Template - Deal data - 2026-11-26.pdf
"""

import argparse  # Import argparse so the script can accept file paths from the command line.
from html.parser import HTMLParser
from copy import deepcopy  # Import copy so selected template formatting objects can be duplicated safely.
from datetime import date  # Import date so the output filename can contain the current date.
from pathlib import Path  # Import Path for reliable Windows path and filename handling.
import re  # Import re to validate six-digit hexadecimal colour values.
import sys  # Import sys so the script can return a non-zero exit code when an error occurs.
from typing import Any, Dict, Iterable, List, Optional, Tuple  # Import typing helpers to document the data structures used by the script.

import yaml  # Import yaml from PyYAML to read the YAML data file.
from pptx import Presentation  # Import Presentation from python-pptx to open and modify the PowerPoint template.
from pptx.enum.text import PP_ALIGN  # Import constants for paragraph alignment.
from pptx.oxml.xmlchemy import OxmlElement  # Import XML helpers for bullet formatting that python-pptx does not expose directly.
from pptx.oxml.ns import qn  # Import the namespace helper used to find PowerPoint XML elements.
from pptx.dml.color import RGBColor  # Import RGBColor to apply explicit text colours.
from pptx.util import Pt  # Import Pt to convert font sizes expressed in points.

PP_SAVE_AS_PDF = 32  # Define the PowerPoint SaveAs file-format value for PDF.
PP_SAVE_AS_OPEN_XML_PRESENTATION = 24  # Define the PowerPoint SaveAs file-format value for a normal PPTX presentation.
HEX_COLOUR_PATTERN = re.compile(r"^#?([0-9A-Fa-f]{6})$")  # Define the accepted six-digit hexadecimal colour format.
ALIGNMENT_MAP = {  # Map readable YAML alignment names to python-pptx alignment constants.
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "centre": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}


def parse_arguments() -> argparse.Namespace:
    """Parse and return command-line arguments supplied by the user."""
    parser = argparse.ArgumentParser(  # Create the command-line parser and describe the script's purpose.
        description="Populate a PowerPoint template from YAML and export it to PDF."
    )
    parser.add_argument("template", type=Path, help="Path to the template .pptx file.")  # Add the required PowerPoint template path argument.
    parser.add_argument("yaml_file", type=Path, help="Path to the .yaml or .yml file.")  # Add the required YAML content path argument.
    parser.add_argument(  # Add an optional output directory while defaulting to the template's directory.
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory. Defaults to the template's directory.",
    )
    return parser.parse_args()  # Return the parsed arguments to the caller.


def validate_input_files(template_path: Path, yaml_path: Path) -> None:
    """Validate that the template and YAML inputs exist and have supported extensions."""
    if not template_path.is_file():  # Stop if the PowerPoint template does not exist as a file.
        raise FileNotFoundError(f"Template file not found: {template_path}")
    if not yaml_path.is_file():  # Stop if the YAML data file does not exist as a file.
        raise FileNotFoundError(f"YAML file not found: {yaml_path}")
    if template_path.suffix.lower() != ".pptx":  # Require the modern PowerPoint Open XML presentation format.
        raise ValueError("The template must be a .pptx file.")
    if yaml_path.suffix.lower() not in {".yaml", ".yml"}:  # Require a conventional YAML filename extension.
        raise ValueError("The content file must have a .yaml or .yml extension.")


def format_yaml_error(yaml_path: Path, yaml_text: str, error: yaml.YAMLError) -> str:
    """Format a PyYAML parsing error with location, source context, and a likely correction."""
    marker = getattr(error, "problem_mark", None) or getattr(error, "context_mark", None)  # Prefer the problem marker because it points at the token that made parsing fail.
    lines = yaml_text.splitlines()  # Split source text into physical lines so the diagnostic can quote the user's YAML.
    if marker is None:  # Fall back to a general message if PyYAML did not provide a source position.
        return f"Could not parse YAML file: {yaml_path}\n\nProblem: {error}\n\nCheck quotes, indentation, colons, and HTML attributes."
    line_no = marker.line + 1  # Convert PyYAML's zero-based indexes to editor-friendly one-based numbers.
    column_no = marker.column + 1
    source_line = lines[marker.line] if 0 <= marker.line < len(lines) else ""  # Read the exact physical line reported by PyYAML.
    first = max(0, marker.line - 2)  # Include two lines before and one after the problem to expose structural YAML mistakes.
    last = min(len(lines), marker.line + 2)
    width = len(str(last))  # Calculate number width so source-line numbers align neatly.
    excerpt = []  # Build the human-readable source excerpt.
    for index in range(first, last):
        prefix = ">" if index == marker.line else " "  # Highlight the failing line with an arrow.
        excerpt.append(f"{prefix} {index + 1:>{width}} | {lines[index]}")  # Append the source line exactly as it appears in the YAML file.
        if index == marker.line:  # Put a caret under PyYAML's reported column.
            excerpt.append(" " * (width + 5 + marker.column) + "^")
    problem = getattr(error, "problem", None) or str(error).splitlines()[0]  # Extract PyYAML's short problem description instead of printing its full technical traceback.
    suggestion = "Check the highlighted line and the line above it for mismatched quotes, indentation, or a missing colon."  # Default to general YAML advice when a specific pattern cannot safely be identified.
    stripped = source_line.lstrip()  # Inspect the failing line for the common HTML-inside-double-quoted-YAML mistake.
    if ":" in source_line and source_line.count(chr(34)) >= 3 and "<" in source_line and ">" in source_line:
        key, raw_value = source_line.split(":", 1)  # Separate the YAML key from its scalar value at the first colon only.
        value = raw_value.strip()  # Trim whitespace around the YAML scalar.
        if value.startswith(chr(34)) and value.endswith(chr(34)):  # Strip only the outer double quotes so HTML attribute quotes remain untouched.
            value = value[1:-1]
        value = value.replace(chr(39), chr(39) * 2)  # Escape single quotes according to YAML single-quoted-scalar syntax.
        replacement = f"{key}: '{value}'"  # Build a ready-to-copy replacement line.
        suggestion = (  # Explain the actual cause and recommend single YAML quotes around HTML-bearing text.
            "This looks like a YAML value enclosed in double quotes that also contains double-quoted HTML attributes. "
            "For example, size=\"14\" can end the YAML string early.\n\n"
            "Recommended change (copy/paste):\n"
            f"    {replacement}\n\n"
            "For longer HTML content, another safe option is a YAML block value using '|'."
        )
    elif "\t" in source_line[: len(source_line) - len(source_line.lstrip())]:  # Give a specific correction when indentation contains a tab.
        suggestion = "The indentation contains a TAB. YAML indentation should use spaces; replace the tab with spaces."
    elif stripped.startswith("-"):  # Give targeted list-indentation advice for a failing YAML list item.
        suggestion = "This is a YAML list item. Align '-' with sibling list items and indent its child fields further."
    return (  # Return a concise diagnostic suitable for direct display by the script's existing ERROR handler.
        "YAML SYNTAX ERROR\n"
        "=================\n"
        f"File:   {yaml_path}\n"
        f"Line:   {line_no}\n"
        f"Column: {column_no}\n"
        f"Problem: {problem}\n\n"
        + "\n".join(excerpt)
        + "\n\nHow to fix it:\n"
        + suggestion
    )


def load_yaml_data(yaml_path: Path) -> Dict[str, Any]:
    """Load YAML safely and turn parser failures into precise, actionable error messages."""
    yaml_text = yaml_path.read_text(encoding="utf-8")  # Read the entire file so parser errors can show the exact source text around the failure.
    try:  # Parse the YAML while intercepting PyYAML's technical scanner/parser exceptions.
        loaded_data = yaml.safe_load(yaml_text)  # Use safe_load so arbitrary Python objects cannot be created via YAML tags.
    except yaml.YAMLError as error:  # Catch all YAML syntax/scanner/parser/composer errors from PyYAML.
        raise ValueError(format_yaml_error(yaml_path, yaml_text, error)) from None  # Raise a regular ValueError so the existing command-line error handler prints only the friendly diagnostic.
    if loaded_data is None:  # Treat an empty YAML document as an empty mapping.
        loaded_data = {}
    if not isinstance(loaded_data, dict):  # Require the YAML document's top level to be an ID-to-content mapping.
        raise ValueError(f"YAML structure error in {yaml_path}: the top level must be a mapping of 'ID: content'.")
    return {str(unique_id): content for unique_id, content in loaded_data.items()}  # Convert every ID to a string while preserving each content definition.


def normalise_colour(value: str) -> RGBColor:
    """Convert '#RRGGBB' or 'RRGGBB' text into a python-pptx RGBColor value."""
    match = HEX_COLOUR_PATTERN.fullmatch(str(value).strip())  # Match the supplied colour against the accepted six-digit format.
    if match is None:  # Reject invalid colour values with a clear message.
        raise ValueError(f"Invalid colour '{value}'. Use a six-digit value such as '#C00000'.")
    hexadecimal_colour = match.group(1)  # Extract the hexadecimal digits without an optional leading hash.
    return RGBColor(  # Convert pairs of hexadecimal digits into red, green, and blue integers.
        int(hexadecimal_colour[0:2], 16),
        int(hexadecimal_colour[2:4], 16),
        int(hexadecimal_colour[4:6], 16),
    )


def remove_bullet_elements(paragraph) -> None:
    """Remove any inherited bullet or no-bullet XML elements from a paragraph."""
    paragraph_properties = paragraph._p.get_or_add_pPr()  # Access the paragraph-properties XML element.
    bullet_element_names = (  # List all bullet-related child element names that may already be present.
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
    for bullet_element_name in bullet_element_names:  # Check each supported bullet-related element name.
        for bullet_element in paragraph_properties.findall(qn(bullet_element_name)):  # Remove every matching child element from the paragraph properties.
            paragraph_properties.remove(bullet_element)


def set_paragraph_bullet(paragraph, enabled: bool, bullet_character: str = "•") -> None:
    """Enable or disable a character bullet for a PowerPoint paragraph."""
    remove_bullet_elements(paragraph)  # Remove inherited bullet declarations to ensure one deterministic setting remains.
    paragraph_properties = paragraph._p.get_or_add_pPr()  # Access the paragraph-properties XML element after cleanup.
    if enabled:  # Add a bullet-character declaration when bullets are enabled.
        bullet_element = OxmlElement("a:buChar")  # Create PowerPoint's XML element for a character bullet.
        bullet_element.set("char", bullet_character or "•")  # Set the bullet character, defaulting to a standard round bullet.
        paragraph_properties.append(bullet_element)  # Append the bullet declaration to the paragraph properties.
    else:  # Add an explicit no-bullet declaration when bullets are disabled.
        no_bullet_element = OxmlElement("a:buNone")  # Create PowerPoint's XML element for no bullet.
        paragraph_properties.append(no_bullet_element)  # Append the no-bullet declaration to the paragraph properties.


def copy_paragraph_defaults(source_paragraph, target_paragraph) -> None:
    """Copy paragraph-level properties from a template paragraph to a target paragraph."""
    source_properties = source_paragraph._p.pPr  # Retrieve the source paragraph-properties element if it exists.
    if source_properties is not None:  # Continue only when the template paragraph contains explicit properties.
        target_properties = target_paragraph._p.pPr  # Retrieve the target paragraph-properties element if it exists.
        if target_properties is not None:  # Remove the target properties before inserting a clean copy.
            target_paragraph._p.remove(target_properties)
        target_paragraph._p.insert(0, deepcopy(source_properties))  # Insert a deep copy so changes do not alter the template XML node.


def copy_run_defaults(source_run, target_run) -> None:
    """Copy run-level character formatting from a template run to a target run."""
    source_properties = source_run._r.rPr  # Retrieve the source run-properties element if it exists.
    if source_properties is not None:  # Continue only when the template run contains explicit character properties.
        target_properties = target_run._r.rPr  # Retrieve the target run-properties element if it exists.
        if target_properties is not None:  # Remove target properties before inserting the copied template properties.
            target_run._r.remove(target_properties)
        target_run._r.insert(0, deepcopy(source_properties))  # Insert a deep copy of the source run properties at the start of the target run.


def apply_run_format(run, run_definition: Dict[str, Any]) -> None:
    """Apply bold, italic, underline, colour, size, and font settings to one text run."""
    if "bold" in run_definition:  # Apply bold only when the YAML explicitly provides a bold value.
        run.font.bold = bool(run_definition["bold"])
    if "italic" in run_definition:  # Apply italic only when the YAML explicitly provides an italic value.
        run.font.italic = bool(run_definition["italic"])
    if "underline" in run_definition:  # Apply underline only when the YAML explicitly provides an underline value.
        run.font.underline = bool(run_definition["underline"])
    if "colour" in run_definition or "color" in run_definition:  # Apply an RGB text colour when the YAML provides a colour value.
        colour_value = run_definition.get("colour", run_definition.get("color"))
        run.font.color.rgb = normalise_colour(str(colour_value))
    if "size" in run_definition:  # Apply font size in points when provided.
        run.font.size = Pt(float(run_definition["size"]))
    if "font" in run_definition:  # Apply a named font family when provided.
        run.font.name = str(run_definition["font"])



class PowerPointHTMLParser(HTMLParser):
    """Convert supported HTML tags into PowerPoint paragraph and run definitions."""

    def __init__(self) -> None:
        """Initialise output, inherited text styles, and list nesting state."""
        super().__init__(convert_charrefs=True)  # Decode entities in text while letting actual HTML tags trigger parser callbacks.
        self.paragraphs: List[Dict[str, Any]] = []  # Store generated paragraph definitions in source order.
        self.current_paragraph: Optional[Dict[str, Any]] = None  # Track the paragraph receiving the next text run.
        self.styles: List[Tuple[str, Dict[str, Any]]] = [("__root__", {})]  # Store inherited formatting for nested tags.
        self.lists: List[str] = []  # Store list nesting so bullet levels can be calculated.

    def _paragraph(self) -> Dict[str, Any]:
        """Return the active paragraph, creating a normal paragraph when needed."""
        if self.current_paragraph is None:  # Create a paragraph if text appears outside an active list item.
            self.current_paragraph = {"runs": []}  # Initialise the paragraph's rich-text run collection.
            self.paragraphs.append(self.current_paragraph)  # Preserve the paragraph's source order.
        return self.current_paragraph  # Return the paragraph that should receive text.

    def _push(self, tag: str, updates: Dict[str, Any]) -> None:
        """Open a formatting tag and inherit the currently active style."""
        style = dict(self.styles[-1][1])  # Copy the parent style so nested formatting is cumulative.
        style.update(updates)  # Apply the new tag's formatting overrides.
        self.styles.append((tag, style))  # Save tag and style for closing-tag validation.

    def _pop(self, tag: str) -> None:
        """Close a formatting tag and validate proper nesting."""
        if len(self.styles) == 1 or self.styles[-1][0] != tag:  # Detect a closing tag that does not match the latest opening tag.
            expected = self.styles[-1][0] if len(self.styles) > 1 else "none"  # Identify the closing tag expected at this point.
            raise ValueError(f"Mismatched HTML closing tag </{tag}>; expected </{expected}>.")  # Report the malformed markup precisely.
        self.styles.pop()  # Restore the parent style after a valid close.

    @staticmethod
    def _size(value: str) -> float:
        """Convert an HTML/CSS size such as 14, 14pt, or 14px into a numeric size."""
        match = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)(?:pt|px)?", value.strip(), re.IGNORECASE)  # Accept common plain and CSS point/pixel forms.
        if match is None:  # Reject unsupported units and non-numeric values.
            raise ValueError(f"Invalid HTML font size '{value}'. Use 14 or 14pt.")  # Explain the accepted syntax.
        return float(match.group(1))  # Return the numeric component for pptx.util.Pt.

    @classmethod
    def _css(cls, value: str) -> Dict[str, Any]:
        """Translate supported inline CSS color and font-size declarations."""
        result: Dict[str, Any] = {}  # Accumulate translated PowerPoint run properties.
        for declaration in value.split(";"):  # Process each CSS declaration independently.
            if not declaration.strip():  # Ignore empty declarations and trailing semicolons.
                continue  # Continue with the next declaration.
            if ":" not in declaration:  # Require standard property:value CSS syntax.
                raise ValueError(f"Invalid CSS declaration '{declaration.strip()}'.")  # Surface malformed CSS rather than ignoring it.
            name, css_value = declaration.split(":", 1)  # Split at the first colon only.
            name, css_value = name.strip().lower(), css_value.strip()  # Normalise property name and value whitespace.
            if name in {"color", "colour"}:  # Support both spelling variants.
                normalise_colour(css_value)  # Validate the RGB colour before rendering.
                result["colour"] = css_value  # Store it using the renderer's existing key.
            elif name == "font-size":  # Support CSS font-size.
                result["size"] = cls._size(css_value)  # Convert it to a numeric point size.
            else:  # Reject CSS the renderer cannot implement.
                raise ValueError(f"Unsupported CSS property '{name}'; use color or font-size.")  # Provide supported alternatives.
        return result  # Return translated formatting.

    def handle_starttag(self, tag: str, attrs) -> None:
        """Interpret supported opening HTML tags."""
        tag = tag.lower()  # Treat tag names case-insensitively.
        attributes = {name.lower(): value for name, value in attrs}  # Convert attributes into a convenient dictionary.
        if tag in {"b", "strong"}:  # Map b and strong to bold.
            self._push(tag, {"bold": True})  # Enable bold until the matching closing tag.
        elif tag == "u":  # Map u to underline.
            self._push(tag, {"underline": True})  # Enable underline until the matching closing tag.
        elif tag == "font":  # Support legacy font colour and size attributes.
            updates: Dict[str, Any] = {}  # Collect formatting declared by the font tag.
            colour = attributes.get("color", attributes.get("colour"))  # Accept color or colour.
            if colour is not None:  # Apply colour only when supplied.
                normalise_colour(colour)  # Validate the six-digit RGB colour.
                updates["colour"] = colour  # Store the colour for the resulting PowerPoint run.
            if attributes.get("size") is not None:  # Apply size only when supplied.
                updates["size"] = self._size(attributes["size"])  # Convert HTML size to a numeric value.
            unsupported = set(attributes) - {"color", "colour", "size"}  # Detect unsupported or misspelled font attributes.
            if unsupported:  # Do not silently discard requested formatting.
                raise ValueError(f"Unsupported <font> attributes: {sorted(unsupported)}")  # Identify rejected attributes.
            self._push(tag, updates)  # Apply the font formatting with normal nesting.
        elif tag == "span":  # Support modern inline CSS.
            unsupported = set(attributes) - {"style"}  # Allow only span's style attribute.
            if unsupported:  # Reject unsupported span attributes.
                raise ValueError(f"Unsupported <span> attributes: {sorted(unsupported)}")  # Identify rejected attributes.
            self._push(tag, self._css(attributes.get("style", "")))  # Translate and apply supported CSS.
        elif tag in {"ul", "ol"}:  # Open a list scope.
            self.lists.append(tag)  # Retain nesting depth and list type.
        elif tag == "li":  # Start one bullet paragraph per list item.
            if not self.lists:  # Require list items to be inside ul or ol.
                raise ValueError("HTML <li> must be inside <ul> or <ol>.")  # Explain the required structure.
            self.current_paragraph = {"runs": [], "bullet": True, "level": len(self.lists) - 1}  # Create a bullet at the current nesting level.
            self.paragraphs.append(self.current_paragraph)  # Add the bullet paragraph to output.
        elif tag in {"br", "p"}:  # Treat br and p starts as paragraph boundaries.
            self.current_paragraph = None  # Force following text into a new PowerPoint paragraph.
        else:  # Reject unsupported markup rather than showing it accidentally.
            raise ValueError(f"Unsupported HTML tag <{tag}>. Supported tags: b, strong, u, font, span, ul, ol, li, br, p. Use &lt; and &gt; to print a tag literally.")  # Provide supported syntax and escaping guidance.

    def handle_endtag(self, tag: str) -> None:
        """Interpret supported closing HTML tags."""
        tag = tag.lower()  # Treat closing tags case-insensitively.
        if tag in {"b", "strong", "u", "font", "span"}:  # Close inline formatting.
            self._pop(tag)  # Restore the parent character style.
        elif tag == "li":  # Finish the active list item.
            self.current_paragraph = None  # Prevent later text from joining the finished bullet.
        elif tag in {"ul", "ol"}:  # Close a list scope.
            if not self.lists or self.lists[-1] != tag:  # Validate list nesting and type.
                raise ValueError(f"Mismatched HTML closing tag </{tag}>.")  # Report malformed list markup.
            self.lists.pop()  # Remove one list nesting level.
            self.current_paragraph = None  # Start a separate paragraph for later ordinary text.
        elif tag in {"br", "p"}:  # Finish a paragraph boundary element.
            self.current_paragraph = None  # Ensure subsequent text starts separately.
        else:  # Reject unsupported closing tags.
            raise ValueError(f"Unsupported HTML closing tag </{tag}>.")  # Identify the unsupported tag.

    def handle_startendtag(self, tag: str, attrs) -> None:
        """Support the self-closing <br/> form."""
        if tag.lower() != "br":  # Permit only the defined self-closing tag.
            raise ValueError(f"Unsupported self-closing HTML tag <{tag}/>.")  # Identify unsupported self-closing markup.
        self.current_paragraph = None  # Begin a new paragraph after <br/>.

    def handle_data(self, data: str) -> None:
        """Add visible text using the currently inherited HTML formatting."""
        if data.strip() == "" and ("\n" in data or "\r" in data):  # Ignore YAML/HTML indentation between block tags.
            return  # Avoid blank paragraphs created only by source formatting.
        if data == "":  # Ignore empty callbacks.
            return  # Nothing needs to be rendered.
        run: Dict[str, Any] = {"text": data}  # Create a run containing the visible text.
        run.update(self.styles[-1][1])  # Apply inherited bold, underline, colour, and size settings.
        self._paragraph()["runs"].append(run)  # Add the formatted run to the active paragraph.

    def validate_complete(self) -> None:
        """Ensure all formatting and list tags were closed."""
        if len(self.styles) != 1:  # Detect unclosed inline formatting.
            raise ValueError(f"Unclosed HTML formatting tag <{self.styles[-1][0]}>.")  # Identify the open formatting tag.
        if self.lists:  # Detect an unclosed list.
            raise ValueError(f"Unclosed HTML list tag <{self.lists[-1]}>.")  # Identify the open list tag.


def parse_html_content(value: str) -> List[Dict[str, Any]]:
    """Parse actual HTML tags while preserving entity-escaped tags as literal visible text."""
    parser = PowerPointHTMLParser()  # Create isolated parser state for this YAML field.
    parser.feed(value)  # Convert actual tags into formatting; entities such as &lt; remain literal text after decoding.
    parser.close()  # Flush buffered parser input.
    parser.validate_complete()  # Reject malformed markup before modifying PowerPoint.
    return parser.paragraphs or [{"runs": [{"text": ""}]}]  # Return an empty paragraph for an empty YAML value.

def value_to_paragraph_definitions(value: Any) -> List[Dict[str, Any]]:
    """Normalise shorthand YAML values into a list of paragraph definitions."""
    if value is None:  # Convert null YAML content into one empty paragraph.
        return [{"runs": [{"text": ""}]}]
    if isinstance(value, str):  # Parse string values so actual HTML tags become PowerPoint formatting.
        return parse_html_content(value)  # Return HTML-derived or ordinary plain-text paragraph definitions.
    if isinstance(value, (int, float, bool)):  # Convert non-string scalar values to visible text without HTML parsing.
        return [{"runs": [{"text": str(value)}]}]  # Return one plain-text paragraph.
    if not isinstance(value, dict):  # Require non-scalar content to be a mapping.
        raise ValueError(f"Unsupported content definition: {value!r}")
    if "paragraphs" in value:  # Use an explicit paragraphs list when one is provided.
        paragraphs = value["paragraphs"]
        if not isinstance(paragraphs, list):  # Require paragraphs to be represented as a YAML list.
            raise ValueError("'paragraphs' must be a list.")
        return paragraphs  # Return the list for later validation and rendering.
    if "runs" in value:  # Convert a top-level runs list into one paragraph.
        return [{key: item for key, item in value.items()}]
    if "text" in value:  # Convert a text mapping into one paragraph while retaining style keys.
        paragraph_definition = dict(value)  # Copy the mapping so the caller's data is not modified.
        run_definition = {"text": paragraph_definition.pop("text")}  # Move the text into a single run while leaving paragraph controls separate.
        run_style_keys = ("bold", "italic", "underline", "colour", "color", "size", "font")  # List style keys that apply to the run rather than the paragraph.
        for style_key in run_style_keys:  # Transfer supplied run-style keys into the run definition.
            if style_key in paragraph_definition:
                run_definition[style_key] = paragraph_definition.pop(style_key)
        paragraph_definition["runs"] = [run_definition]  # Store the run in the normalised paragraph definition.
        return [paragraph_definition]  # Return the one-paragraph list.
    raise ValueError("Content mappings must contain 'text', 'runs', or 'paragraphs'.")  # Reject mappings that do not contain a recognised content structure.


def populate_text_frame(text_frame, content_definition: Any) -> None:
    """Replace a text frame with YAML-defined paragraphs, bullets, and rich-text runs."""
    source_paragraph = text_frame.paragraphs[0] if text_frame.paragraphs else None  # Preserve the first template paragraph as a source of paragraph defaults.
    source_run = source_paragraph.runs[0] if source_paragraph and source_paragraph.runs else None  # Preserve the first template run as a source of character defaults.
    paragraph_definitions = value_to_paragraph_definitions(content_definition)  # Convert shorthand YAML into a consistent sequence of paragraph definitions.
    text_frame.clear()  # Clear all current text while retaining the shape and its text-frame properties.
    for paragraph_index, paragraph_definition in enumerate(paragraph_definitions):  # Render each YAML paragraph in order.
        if not isinstance(paragraph_definition, dict):  # Require every paragraph definition to be a mapping.
            raise ValueError("Each paragraph definition must be a mapping.")
        if paragraph_index == 0:  # Reuse the paragraph created by clear() for the first item.
            paragraph = text_frame.paragraphs[0]
        else:  # Add a new paragraph for every later item.
            paragraph = text_frame.add_paragraph()
        if source_paragraph is not None:  # Restore template paragraph properties before applying YAML overrides.
            copy_paragraph_defaults(source_paragraph, paragraph)
        paragraph.level = int(paragraph_definition.get("level", 0))  # Read the paragraph's indentation level, defaulting to level zero.
        if "align" in paragraph_definition:  # Apply paragraph alignment when supplied.
            alignment_name = str(paragraph_definition["align"]).strip().lower()
            if alignment_name not in ALIGNMENT_MAP:
                raise ValueError(f"Unsupported alignment '{alignment_name}'.")
            paragraph.alignment = ALIGNMENT_MAP[alignment_name]
        if "bullet" in paragraph_definition:  # Enable or disable bullet formatting when explicitly supplied.
            set_paragraph_bullet(
                paragraph,
                bool(paragraph_definition["bullet"]),
                str(paragraph_definition.get("bullet_char", "•")),
            )
        run_definitions = paragraph_definition.get("runs", [{"text": ""}])  # Read the run list, or create one empty run when omitted.
        if not isinstance(run_definitions, list):  # Require runs to be represented as a YAML list.
            raise ValueError("'runs' must be a list.")
        for run_definition in run_definitions:  # Add each rich-text run to the paragraph.
            if isinstance(run_definition, (str, int, float, bool)):  # Allow a scalar run as shorthand for plain text.
                run_definition = {"text": str(run_definition)}
            if not isinstance(run_definition, dict):  # Require a mapping after shorthand conversion.
                raise ValueError("Each run definition must be text or a mapping.")
            run = paragraph.add_run()  # Add a new PowerPoint run to the current paragraph.
            if source_run is not None:  # Restore template character formatting before applying YAML overrides.
                copy_run_defaults(source_run, run)
            raw_text = run_definition.get("text", "")  # Set the run's visible text, treating null as an empty string.
            run.text = "" if raw_text is None else str(raw_text)
            apply_run_format(run, run_definition)  # Apply every supplied rich-text formatting override.


def iter_shapes(shapes: Iterable[Any]) -> Iterable[Any]:
    """Yield every shape, recursively including shapes inside PowerPoint groups."""
    for shape in shapes:  # Examine each shape in the supplied shape collection.
        yield shape  # Yield the current shape to the caller.
        if hasattr(shape, "shapes"):  # Recurse when the shape exposes a nested shapes collection, as group shapes do.
            yield from iter_shapes(shape.shapes)


def find_targets(presentation: Presentation, unique_id: str) -> List[Tuple[int, Any]]:
    """Find text shapes whose Selection Pane name equals the ID or whose text contains {{ID}}."""
    token = "{{" + unique_id + "}}"  # Build the token form used by the fallback matching convention.
    matches: List[Tuple[int, Any]] = []  # Create the list that will hold slide numbers and matching shapes.
    for slide_number, slide in enumerate(presentation.slides, start=1):  # Search every slide in the presentation.
        for shape in iter_shapes(slide.shapes):  # Search every normal and grouped shape on the slide.
            if not getattr(shape, "has_text_frame", False):  # Skip shapes that cannot contain editable text.
                continue
            if shape.name == unique_id or token in shape.text:  # Match the preferred Selection Pane name or the fallback text token.
                matches.append((slide_number, shape))
    return matches  # Return every target so one YAML ID may intentionally populate several shapes.


def populate_presentation(template_path: Path, yaml_data: Dict[str, Any], output_pptx: Path) -> None:
    """Populate all YAML IDs in the template and save the resulting PPTX file."""
    presentation = Presentation(str(template_path))  # Open the PowerPoint template.
    missing_ids: List[str] = []  # Collect IDs that cannot be found so all problems can be reported together.
    for unique_id, content_definition in yaml_data.items():  # Process each YAML ID and its content definition.
        targets = find_targets(presentation, unique_id)  # Locate all template text fields associated with this ID.
        if not targets:  # Record the ID and continue when the template contains no matching field.
            missing_ids.append(unique_id)
            continue
        for _slide_number, shape in targets:  # Populate every matching field, including repeated IDs on different slides.
            populate_text_frame(shape.text_frame, content_definition)
    if missing_ids:  # Stop before saving if one or more YAML IDs were unmatched.
        missing_list = ", ".join(sorted(missing_ids))
        raise ValueError(f"These YAML IDs were not found in the template: {missing_list}")
    presentation.save(str(output_pptx))  # Save the populated presentation as a new PPTX file.


def export_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> None:
    """Use the installed Windows PowerPoint application to export a PPTX file to PDF."""
    try:  # Import COM support only at export time so earlier validation errors stay clear.
        import win32com.client
    except ImportError as exc:  # Convert a missing pywin32 installation into a targeted instruction.
        raise RuntimeError(
            "pywin32 is required for PDF export. Install it with: "
            "python -m pip install pywin32"
        ) from exc
    powerpoint_application = None  # Create variables in advance so the finally block can close partially opened objects.
    powerpoint_presentation = None
    try:  # Use COM automation inside a cleanup-protected block.
        powerpoint_application = win32com.client.DispatchEx("PowerPoint.Application")  # Launch a separate PowerPoint automation instance.
        powerpoint_presentation = powerpoint_application.Presentations.Open(  # Keep the PowerPoint application hidden while the conversion runs. Open the populated PPTX as read-only and without displaying a window.
            str(pptx_path.resolve()),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        powerpoint_presentation.SaveAs(str(pdf_path.resolve()), PP_SAVE_AS_PDF)  # Save the presentation as PDF using PowerPoint's PDF file-format constant.
    finally:  # Always release the PowerPoint presentation and application objects.
        if powerpoint_presentation is not None:  # Close the open presentation if it was successfully created.
            powerpoint_presentation.Close()
        if powerpoint_application is not None:  # Quit the separate PowerPoint instance if it was successfully created.
            powerpoint_application.Quit()


def build_output_paths(template_path: Path, yaml_path: Path, output_dir: Path) -> Tuple[Path, Path]:
    """Build output PPTX and PDF paths using template name, YAML name, and today's date."""
    template_name = template_path.stem  # Remove the extensions from the template and YAML filenames.
    yaml_name = yaml_path.stem
    
    date_text = date.today().isoformat()  # Format the current local date as ISO YYYY-MM-DD.
    output_stem = f"{template_name} - {yaml_name} - {date_text}"  # Construct the filename stem requested by the user.
    return output_dir / f"{output_stem}.pptx", output_dir / f"{output_stem}.pdf"  # Return the populated PPTX path and its corresponding PDF path.


def main() -> int:
    """Coordinate validation, template population, PPTX saving, and PDF export."""
    arguments = parse_arguments()  # Parse the paths and options supplied on the command line.
    template_path = arguments.template.expanduser().resolve()  # Resolve the template path to an absolute filesystem location.
    yaml_path = arguments.yaml_file.expanduser().resolve()  # Resolve the YAML path to an absolute filesystem location.
    validate_input_files(template_path, yaml_path)  # Validate both input paths before opening either file.
    output_dir = (  # Use the requested output directory or default to the template's directory.
        arguments.output_dir.expanduser().resolve()
        if arguments.output_dir is not None
        else template_path.parent
    )
    output_dir.mkdir(parents=True, exist_ok=True)  # Create the output directory, including missing parent directories.
    yaml_data = load_yaml_data(yaml_path)  # Load and validate the YAML ID-to-content mapping.
    output_pptx, output_pdf = build_output_paths(template_path, yaml_path, output_dir)  # Build the dated output filenames.
    populate_presentation(template_path, yaml_data, output_pptx)  # Populate the PowerPoint template and save the intermediate deliverable PPTX.
    export_pptx_to_pdf(output_pptx, output_pdf)  # Export the populated PPTX to PDF using the installed PowerPoint application.
    print(f"Created PPTX: {output_pptx}")  # Print both output paths so batch jobs can capture them.
    print(f"Created PDF:  {output_pdf}")
    return 0  # Return zero to indicate successful completion.


if __name__ == "__main__":  # Run main() only when this file is executed directly rather than imported as a module.
    try:  # Convert expected runtime failures into a concise command-line error message.
        sys.exit(main())
    except Exception as error:  # Catch ordinary exceptions while preserving a non-zero exit code for automation.
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)


