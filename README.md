<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/logo-dark.svg">
    <img src="assets/logo.svg" alt="ppt-template-to-pdf" width="132">
  </picture>
</p>

# ppt-template-to-pdf

`ppt_template_to_pdf.py` fills a PowerPoint template with data from a YAML file. The script saves the filled deck as `.pptx` and exports the same deck to `.pdf`.

## What the script does

- Opens a template with `python-pptx`.
- Reads a YAML file that maps field IDs to content.
- Finds every matching text box in the template and replaces its text.
- Keeps the template's paragraph and character formatting as the default look.
- Saves a new `.pptx` file. The template file is never modified.
- Exports the new `.pptx` to PDF with the installed desktop PowerPoint. Use `--no-pdf` to skip this step.

Output names use the template name, the YAML name, and today's date:

```
Template - Deal data - 2026-11-26.pptx
Template - Deal data - 2026-11-26.pdf
```

If those files already exist, the script adds a number instead of overwriting them:

```
Template - Deal data - 2026-11-26 (2).pptx
```

Pass `--overwrite` to replace the earlier files instead.

## Requirements

- Python 3.10 or later.
- `python-pptx` and `PyYAML`.
- PDF export: Windows, a desktop install of Microsoft PowerPoint, and `pywin32`. The Microsoft Store, web, and Mac versions of PowerPoint do not expose the COM interface. Linux and macOS work with `--no-pdf`.

## Install dependencies

```
python -m pip install -r requirements.txt
```

Or install the three packages directly:

```
python -m pip install python-pptx PyYAML pywin32
```

`requirements.txt` marks `pywin32` as Windows-only, so the same file works on Linux and macOS.

Known working versions:

| Package | Version |
| --- | --- |
| python-pptx | 1.0.2 |
| PyYAML | 6.0.4 |
| pywin32 | 312 |
| lxml | 6.1.3 |
| pillow | 13.0.0 |

If `pywin32` installs but COM calls fail with `Invalid class string`, run the post-install script once:

```
python "C:\path\to\python\Scripts\pywin32_postinstall.py" -install
```

Find the real path with `python -c "import sys; print(sys.prefix)"`.

Install the script itself to get a `ppt-fill` command:

```
python -m pip install .
ppt-fill "Template.pptx" "Deal data.yaml"
```

## Use

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"
```

Write outputs to another folder:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --output-dir "Output"
```

Flags:

| Flag | Effect |
| --- | --- |
| `-o`, `--output-dir DIR` | Output folder. Created if missing. Defaults to the template's folder. |
| `--no-pdf` | Create only the PPTX. Works without PowerPoint and without Windows. |
| `--dry-run` | Check the YAML and the template, print the matches, write nothing. |
| `--list-fields` | Print every text shape name and token in the template, then exit. Needs no YAML file. |
| `--overwrite` | Replace existing output files instead of adding `(2)`, `(3)`, and so on. |
| `-v`, `--verbose` | Print every field that is filled, and every text shape that is left alone. |
| `-q`, `--quiet` | Print errors only. |
| `--version` | Print the version and exit. |

Quote paths that contain spaces.

Exit codes:

| Code | Meaning |
| --- | --- |
| 0 | Success. |
| 1 | Unexpected failure. |
| 2 | Bad arguments, missing files, or invalid YAML. |
| 3 | The YAML refers to IDs the template does not contain. |
| 4 | PDF export failed. The PPTX may still exist. |

## Check the template first

`--list-fields` prints every text shape, its slide, and its current text. It also shows which IDs are already written as `{{tokens}}`.

```
$ python ppt_template_to_pdf.py "Template.pptx" --list-fields
Slide  Shape name         Match by       Text
------------------------  -------------  --------------------
1      CLIENT_NAME        shape name     {{CLIENT_NAME}}
1      TRANSACTION_TITLE  shape name     Project ___ Financing
3      TextBox 7          token SUMMARY  {{SUMMARY}}
```

Use those names, or the IDs inside the braces, as top-level YAML keys.

Preview a real run without writing files:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --dry-run -v
```

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

Matching rules:

- Shape names match exactly first.
- Shape names then match case-insensitively, with a warning.
- Tokens match with or without inner spaces, and are case-insensitive.
- Shapes inside groups are searched.
- One ID can match several shapes. Every match gets the same content.
- If one shape is targeted by two IDs, the script warns and the later ID wins.

Tables and pictures have no text frame, so the script skips them. Put the ID on a real text box.

Templates can be `.pptx`, `.pptm`, or `.potx`. The output is always `.pptx`, so macros are dropped.

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
    - number: true
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
- `bold`, `italic`, `underline`: `true` or `false`. The strings `"true"`, `"false"`, `"yes"`, `"no"`, `"on"`, and `"off"` also work.
- `colour` or `color`: six hex digits, with or without `#`, for example `#C00000`.
- `size`: font size in points, between 1 and 400.
- `font`: font family name, for example `Arial`.

Paragraph fields:

- `runs`: list of runs. A run can also be a plain string.
- `text`: shorthand for one run. Style keys such as `bold` move to that run.
- `bullet`: `true` or `false`.
- `bullet_char`: bullet character. Default is `•`. Must be one character.
- `number`: numbered list marker instead of a bullet.
- `level`: indentation level, `0` to `8`.
- `align`: `left`, `center`, `centre`, `right`, or `justify`.

Paragraph and run formatting that the YAML does not set comes from the first paragraph and first run of the template text box.

Unknown keys are rejected rather than ignored. A typo like `colourr` reports `did you mean 'colour'?`.

## HTML inside YAML strings

Supported tags:

- `b`, `strong`: bold.
- `u`: underline.
- `font`: `color` and `size` attributes.
- `span`: inline CSS with `color` and `font-size` only.
- `ul`: bulleted list.
- `ol`: numbered list.
- `li`: one paragraph in the current list.
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

STEPS: |
  <ol>
    <li>First step</li>
    <li>Second step</li>
  </ol>
```

To print a tag as text, escape the angle brackets:

```yaml
LITERAL_TAG: "&lt;b&gt;This is not bold&lt;/b&gt;"
```

Use single quotes around a YAML value that contains HTML with double-quoted attributes. Double-quoted YAML strings end at the first inner `"`. The alternative is a block value with `|`.

## What can go wrong

| Message or symptom | Cause | Fix |
| --- | --- | --- |
| `Template file not found` | Wrong path or missing file. | Check the path. Quote paths with spaces. |
| `YAML file not found` | Wrong path. | Check the path and the `.yaml` / `.yml` extension. |
| `Template must be one of .potx, .pptm, .pptx` | The file is an old `.ppt`. | Open the file in PowerPoint and save it as `.pptx`. |
| `Data file must be .yaml or .yml` | The data file has another extension. | Rename the file or fix the extension. |
| `Output path is a file, not a folder` | `-o` points at a file. | Pass a folder path, or delete the file. |
| `YAML COMMENT ERROR` | A double-quoted value contains `colour="#C00000"`. YAML reads `#` as a comment and cuts the value there. | The message prints the fixed line. Wrap the value in single quotes or use a `\|` block. |
| `YAML SYNTAX ERROR` with line and column | Bad quotes, tabs, or indentation. | Read the quoted line and the suggestion under it. Use spaces, never tabs. |
| `found duplicate key` | The same ID appears twice. | Remove the duplicate or rename one of them. |
| `ID ... is not text` | A key such as `NO`, `ON`, or `YES` was read as a boolean, not as a word. | Quote the key: `'NO': content`. |
| `is empty` or `contains no IDs` | The YAML file has no entries. | Add at least one `ID: content` entry. |
| `is not valid UTF-8 text` | The file was saved in a legacy code page. | Save the YAML file as UTF-8. A UTF-8 BOM is fine. |
| `the top level must be a mapping` | The YAML file is a list or a single value. | Start the file with `ID: value` lines at column zero. |
| `These YAML IDs were not found in the template` | Shape name differs from the YAML key. The field is a table or a picture. | The message suggests close matches. Run `--list-fields` to copy the exact name. |
| `unknown content key` or `unknown run key` | A key is misspelled or not supported. | Use a suggested key. The message lists every allowed key. |
| `invalid colour` | Colour is not six hex digits. | Use `#C00000`. Colour names such as `red` do not work. |
| `unsupported HTML tag` | The tag is not in the supported list. | Use a supported tag. Escape the angle brackets to print a tag as text. |
| `mismatched HTML closing tag` or `unclosed HTML` | Tags are not closed or are nested in the wrong order. | Close every tag, inside out: `<b><u>text</u></b>`. |
| `unsupported CSS property` | `span style` uses a property other than `color` or `font-size`. | Remove the property, or use a `font` tag or run fields. |
| `use 'bullet' or 'number', not both` | One paragraph asks for a bullet and a number. | Keep one of the two. |
| `'bullet_char' must be a single character` | The bullet character is a longer string. | Use one character. |
| `unsupported alignment` | `align` value is not in the list. | Use `left`, `center`, `centre`, `right`, or `justify`. |
| `PDF export needs Windows` | The script ran on Linux or macOS. | Add `--no-pdf`, or run on Windows. |
| `PDF export needs pywin32` | `pywin32` is missing from the active environment. | Install it with `python -m pip install pywin32`. Check the interpreter is the same one, for example inside a venv. |
| `PowerPoint automation is not registered` | COM cannot start PowerPoint. Usually a Microsoft Store install or broken `pywin32` registration. | Install desktop PowerPoint. Run the `pywin32_postinstall.py` script shown above. |
| `PowerPoint could not open the file` | The PPTX is open, read-only, or in a protected or blocked location. | Close the file. Move it out of OneDrive or a network share. Unblock it in file Properties. |
| PPTX created but no PDF | Export failed, or PowerPoint is unavailable. | Read the error text. Fix PowerPoint, then run again. |
| `PDF export failed after 3 attempts` | PowerPoint kept rejecting the call. | Close every open Office window and run again. |
| `Cannot write ...` | The output file is open, or the folder is read-only. | Close the file. Write to a folder you own. |
| Output path is long warning | Windows path limit. | Use a short `-o` folder such as `C:\Out`. |
| Text overflows the box | The template box is too small. | Enlarge the box or enable autofit in the template. |
| PDF fonts differ from the template | A template font is not installed on this machine. | Install the font, or save the template with an available font. |

## Exit codes in scripts

The script never prints a traceback for expected problems. It prints `ERROR: ...` to stderr and returns a code from the table above. Check `$LASTEXITCODE` in PowerShell and `$?` in Bash.

## Run the tests

```
python -m pip install -r requirements-dev.txt
python -m pytest
```

The tests cover YAML loading, the HTML parser, matching, and PPTX writing. They do not need PowerPoint. PDF export is checked by hand.

## License

MIT. See [LICENSE](LICENSE).
