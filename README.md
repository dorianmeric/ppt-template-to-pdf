# ppt-template-to-pdf

`ppt_template_to_pdf.py` fills a PowerPoint template with data from a YAML file. The script saves the filled deck as `.pptx` and exports the same deck to `.pdf`.

## What the script does

- Opens a `.pptx` template with `python-pptx`.
- Reads a YAML file that maps field IDs to content.
- Finds each matching text box in the template and replaces its text.
- Keeps the template's own paragraph and character formatting as the default look.
- Saves a new `.pptx` file. The template file is never modified.
- Exports the new `.pptx` to PDF with the installed desktop PowerPoint.

Output names use the template name, the YAML name, and today's date:

```
Template - Deal data - 2026-11-26.pptx
Template - Deal data - 2026-11-26.pdf
```

## Requirements

- Windows. PDF export uses PowerPoint COM automation, so Linux and macOS do not work.
- A desktop install of Microsoft PowerPoint. The Microsoft Store, web, and Mac versions do not expose the COM interface.
- Python 3.10 or later.
- Three packages: `python-pptx`, `PyYAML`, `pywin32`.

## Install dependencies

```
python -m pip install python-pptx PyYAML pywin32
```

Known working versions:

| Package | Version |
| --- | --- |
| python-pptx | 1.0.2 |
| PyYAML | 6.0.4 |
| pywin32 | 312 |
| lxml | 6.1.3 |
| pillow | 13.0.0 |
| xlsxwriter | 3.3.4 |

Pin the versions if you need a reproducible setup:

```
python -m pip install "python-pptx==1.0.2" "PyYAML==6.0.4" "pywin32==312"
```

If `pywin32` installs but COM calls fail with `Invalid class string`, run the post-install script once:

```
python "C:\path\to\python\Scripts\pywin32_postinstall.py" -install
```

Find the real path with `python -c "import sys; print(sys.prefix)"`.

## Use

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"
```

Write outputs to another folder:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --output-dir "Output"
```

- First argument: path to the `.pptx` template.
- Second argument: path to the `.yaml` or `.yml` data file.
- `--output-dir`: optional output folder. Missing folders are created. Default is the template's folder.

Quote paths that contain spaces.

Exit code `0` means success. Exit code `1` means the script printed an error to stderr.

Running the script twice on the same day overwrites the earlier files, because the date part of the filename does not change.

## Prepare the template

Use one of two methods.

**Method 1 (recommended): shape names.**

- Open the template in PowerPoint.
- Go to Home → Arrange → Selection Pane.
- Rename each text box or placeholder to the field ID.
- Use the same ID as the top-level YAML key.

```
CLIENT_NAME
TRANSACTION_TITLE
KEY_POINTS
```

**Method 2: tokens in the text.**

- Type a token such as `{{CLIENT_NAME}}` inside a text box.
- Use the ID inside the braces as the top-level YAML key.

The script checks shape names first, then tokens. Shape names are case sensitive. Shapes inside groups are searched too. One ID can match several shapes, and every match gets the same content.

Tables and pictures have no text frame, so the script skips them. Put the ID on a real text box.

## YAML format

The top level must be a mapping of `ID: content`.

A value can be:

- A plain string. Actual HTML tags in the string are converted to formatting.
- A number or boolean. Printed as plain text.
- A mapping with `text`, `runs`, or `paragraphs`.

Full example:

```yaml
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
```

Run fields:

- `text`: the visible text.
- `bold`, `italic`, `underline`: `true` or `false`.
- `colour` or `color`: six hex digits, with or without `#`, for example `#C00000`.
- `size`: font size in points.
- `font`: font family name, for example `Arial`.

Paragraph fields:

- `runs`: list of runs. A run can also be a plain string.
- `text`: shorthand for one run. Style keys such as `bold` move to that run.
- `bullet`: `true` or `false`.
- `bullet_char`: bullet character. Default is `•`.
- `level`: indentation level, starting at `0`.
- `align`: `left`, `center`, `centre`, `right`, or `justify`.

Paragraph and run formatting that the YAML does not set comes from the first paragraph and first run of the template text box.

## HTML inside YAML strings

Supported tags:

- `b`, `strong`: bold.
- `u`: underline.
- `font`: `color` and `size` attributes.
- `span`: inline CSS with `color` and `font-size` only.
- `ul`, `ol`: list scopes.
- `li`: one bullet paragraph. Nesting sets the paragraph level.
- `br`, `p`: paragraph breaks.

Examples:

```yaml
TITLE: "<b>Bold text</b> and <u>underlined text</u>"
COLOURED: '<font color="#C00000">Red text</font>'
SIZED: '<font size="14">14-point text</font>'
COMBINED: '<span style="color:#C00000; font-size:14pt"><b>Red, bold, 14-point text</b></span>'

BULLETS: |
  <ul>
    <li>First bullet</li>
    <li>Second bullet with <b>bold text</b></li>
  </ul>
```

To print a tag as text, escape the angle brackets:

```yaml
LITERAL_TAG: "&lt;b&gt;This is not bold&lt;/b&gt;"
```

Use single quotes around a YAML value that contains HTML with double-quoted attributes. Double-quoted YAML strings end at the first inner `"`. The alternative is a block value with `|`.

## What can go wrong

| Message or symptom | Cause | Fix |
| --- | --- | --- |
| `Template file not found` | Wrong path, missing quotes, or missing file extension. | Check the path. Quote paths with spaces. Pass the full name, including `.pptx`. |
| `YAML file not found` | Wrong path. | Check the path and the `.yaml` / `.yml` extension. |
| `The template must be a .pptx file.` | The file is an old `.ppt`. | Open the file in PowerPoint and save it as `.pptx`. |
| `The content file must have a .yaml or .yml extension.` | The data file is `.txt` or another extension. | Rename the file or fix the extension. |
| `YAML SYNTAX ERROR` with line and column | Bad quotes, tabs, or indentation. Most common cause: double-quoted YAML value holding double-quoted HTML attributes such as `size="14"`. | Read the quoted line and the suggestion below it. Switch to single quotes or a `|` block. Use spaces, never tabs. |
| `the top level must be a mapping of 'ID: content'` | The YAML file is a list, a single string, or empty of keys. | Start the file with `ID: value` lines at column zero. |
| `These YAML IDs were not found in the template: ...` | Shape name differs from the YAML key in case, spaces, or underscores. The field is a table or picture. The token is misspelled. | Copy the name from the Selection Pane into the YAML. Use a text box for the field. Check the `{{TOKEN}}` spelling. |
| `Invalid colour '...'` | Colour is not six hex digits. | Use `#C00000`. Colour names such as `red` do not work. |
| `Unsupported HTML tag <...>` | The tag is not in the supported list. | Use a supported tag. Escape the angle brackets to print the tag as text. |
| `Mismatched HTML closing tag` or `Unclosed HTML formatting tag` | Tags are not closed or are nested in the wrong order. | Close every tag. Keep nesting inside out: `<b><u>text</u></b>`. |
| `Unsupported CSS property` | `span style` uses a property other than `color` or `font-size`. | Remove the property, or use a `font` tag or run fields instead. |
| `Unsupported alignment '...'` | `align` value is not in the list. | Use `left`, `center`, `centre`, `right`, or `justify`. |
| `pywin32 is required for PDF export` | `pywin32` is missing from the active Python environment. | Run `python -m pip install pywin32`. Check you used the same interpreter, for example inside a venv. |
| `Invalid class string` or `Class not registered` | COM cannot start PowerPoint. Usually a Microsoft Store install, a broken `pywin32` registration, or no desktop PowerPoint. | Install Microsoft Office desktop with PowerPoint. Run the `pywin32_postinstall.py` script shown above. |
| `PowerPoint could not open the file` or a COM error during export | The PPTX is already open in PowerPoint, is read-only, or sits in a protected or blocked location. | Close the file in PowerPoint. Move the file out of OneDrive or a network share. Right-click the file, choose Properties, and Unblock it. |
| The PPTX is created but no PDF appears | PDF export failed or PowerPoint is not available. | Read the error text. Fix the PowerPoint issue, then run the script again. |
| `Permission denied` on save | The output file is open in PowerPoint or Excel, or the folder is read-only. | Close the file. Write to a folder you own. |
| Filenames too long or save fails on a deep folder | Windows path limit. | Use a short `--output-dir` such as `C:\Out`. |
| Text overflows the box | Template box is too small for the content. | Enlarge the box in the template. Enable autofit on the template placeholder. |
| PDF fonts differ from the template | A font used in the template is not installed on the machine. | Install the font. Save the template with an available font. |
| Accented or non-Latin characters look wrong | The YAML file is not UTF-8. | Save the YAML file as UTF-8. |

## License

MIT. See [LICENSE](LICENSE).
