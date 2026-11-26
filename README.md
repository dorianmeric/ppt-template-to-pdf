<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/logo-dark.svg">
    <img src="assets/logo.svg" alt="ppt-template-to-pdf" width="132">
  </picture>
</p>

# Powerpoint-Template-to-PDF

**A Python script to update a PowerPoint template with variables — fill a `.pptx` deck from a YAML or JSON file, then export it to PDF.**

`ppt_template_to_pdf.py` takes a PowerPoint template, replaces the placeholders with values from a **YAML or JSON** data file, saves the filled deck as `.pptx`, and exports the same deck to `.pdf`. It is a command-line tool and a Python module, so it works by hand, from a script, from a scheduled job, or in CI.

On Windows, macOS, and Linux you do not need Python at all: grab the matching binary from [Releases](../../releases) — one file with the runtime and every dependency inside. See [Pre-built executables](#pre-built-executables).

People find this repo looking for:

- a **script to update PowerPoint with variables**
- a **PowerPoint template updated dynamically** at runtime
- **how to update PowerPoint in Python** without losing the template's design
- how to do a **mail merge for PowerPoint** (`{{PLACEHOLDER}}` → real values)
- how to **fill a PPTX template from JSON, YAML, Excel, or a database**
- how to **update slide master and layout placeholders** in a PowerPoint template
- how to **replace the same placeholder everywhere** in a deck, not just on the first slide
- how to **set the PDF title and metadata** of an exported deck
- how to **generate PowerPoint slides programmatically** and ship them as PDF

That is exactly what this does. No PowerPoint add-in, no VBA, no Office 365 subscription on the build machine (except for the optional PDF/legacy-file step).

**Quick start**

```
python -m pip install -r requirements.txt
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"
```

JSON works the same way, just with another extension:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.json"
```

Output:

```
Template - Deal data - 2026-11-26.pptx
Template - Deal data - 2026-11-26.pdf
```

## Contents

- [Powerpoint-Template-to-PDF](#powerpoint-template-to-pdf)
  - [Contents](#contents)
  - [What the script does](#what-the-script-does)
  - [Requirements](#requirements)
  - [Install](#install)
  - [Use](#use)
    - [What a run prints](#what-a-run-prints)
  - [Check the template first](#check-the-template-first)
  - [Prepare the template](#prepare-the-template)
    - [Method 1 (recommended): shape names](#method-1-recommended-shape-names)
    - [Method 2: tokens in the text](#method-2-tokens-in-the-text)
    - [Method 3: on a slide layout or the slide master](#method-3-on-a-slide-layout-or-the-slide-master)
    - [Matching rules](#matching-rules)
  - [HTML inside YAML or JSON values](#html-inside-yaml-or-json-values)
  - [Pure YAML format](#pure-yaml-format)
  - [JSON data files](#json-data-files)
    - [Converting between YAML and JSON](#converting-between-yaml-and-json)
  - [Document properties](#document-properties)
  - [Use it from Python](#use-it-from-python)
  - [Troubleshooting](#troubleshooting)
  - [FAQ: update PowerPoint with Python](#faq-update-powerpoint-with-python)
  - [How this compares to other tools](#how-this-compares-to-other-tools)
  - [Search terms](#search-terms)
  - [Pre-built executables](#pre-built-executables)
  - [Releases](#releases)
  - [Run the tests](#run-the-tests)
  - [License](#license)

## What the script does

- Opens a template with `python-pptx`.
- Reads a YAML or JSON file that maps field IDs to content.
- Searches every text shape on every slide, every slide layout, and every slide master, including shapes inside groups.
- Finds **every** matching text box and replaces its text, so a field that appears five times is filled five times, not only at the first hit.
- Keeps the template's paragraph and character formatting as the default look.
- Saves a new `.pptx` file. The template file is never modified.
- Sets document properties (title, subject, author, keywords, dates) from a `PROPERTIES` block, which is what PowerPoint and the exported PDF show as metadata.
- Converts a legacy `.ppt` or `.pot` template to `.pptx` first, with the installed PowerPoint.
- Exports the new `.pptx` to PDF with the installed desktop PowerPoint. Use `--no-pdf` to skip this step.

In other words: **one template, one data file (YAML or JSON), any number of decks.** The template keeps the branding, the master slides, the fonts and the layout; only the text changes.

Output names use the template name, the data file name, and today's date:

```
Template - Deal data - 2026-11-26.pptx
Template - Deal data - 2026-11-26.pdf
```

Pass `-n` / `--name` to name both files yourself:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --name "Q3 update"
```

```
Q3 update.pptx
Q3 update.pdf
```

The extension is added for you, so `--name "Q3 update.pdf"` produces the same two files. A name that cannot be used as a file name is rejected with an error instead of failing later:

| Name | Result |
| --- | --- |
| `Q3 update`, `Q3_update`, `v1.2 report` | Accepted. |
| `Q3 update.pdf`, `Q3 update.pptx` | Accepted. The extension is dropped and both files are named `Q3 update`. |
| `out/report` | Rejected. No folders in the name; use `--output-dir` for that. |
| `Q3: update`, `Q3?`, `a*b`, `a|b`, `a"b` | Rejected. Characters Windows forbids in a file name. |
| `CON`, `NUL`, `COM1`, `LPT1` | Rejected. Names Windows reserves for devices. |
| `.hidden`, `trailing.` | Rejected. A name must not start or end with a dot. |

If those files already exist, the script adds a number instead of overwriting them:

```
Template - Deal data - 2026-11-26 (2).pptx
Q3 update (2).pptx
```

Pass `--overwrite` to replace the earlier files instead.

## Requirements

- Python 3.10 or later.
- `python-pptx` and `PyYAML`. JSON needs no extra package.
- Legacy `.ppt` and `.pot` templates: Windows, PowerPoint, and `pywin32`. `python-pptx` cannot read the old binary format, so the script asks PowerPoint to convert the file first. This applies even with `--no-pdf`.
- PDF export: Windows, a desktop install of Microsoft PowerPoint, and `pywin32`. The Microsoft Store, web, and Mac versions of PowerPoint do not expose the COM interface. Linux and macOS work with `--no-pdf`.

## Install

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

Install the script itself to get a `ppt-template-to-pdf` command:

```
python -m pip install .
ppt-template-to-pdf "Template.pptx" "Deal data.yaml"
```

## Use

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"
python ppt_template_to_pdf.py "Template.pptx" "Deal data.json"
```

The data file extension picks the format: `.yaml` and `.yml` are read as YAML, `.json` is read as JSON. Everything else — the IDs, the values, the HTML formatting, the flags — behaves the same.

Write outputs to another folder, and choose the base name of both files:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --output-dir "Output"
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" -o "Output" -n "Q3 update"
```

Flags:

| Flag | Effect |
| --- | --- |
| `-o`, `--output-dir DIR` | Output folder. Created if missing. Defaults to the template's folder. |
| `-n`, `--name NAME` | Base name for both output files, without extension. Defaults to `<template> - <data> - <date>`. A `.pptx` or `.pdf` suffix is dropped. |
| `--no-pdf` | Create only the PPTX. Works without PowerPoint and without Windows. |
| `--dry-run` | Check the data file and the template, print the matches, write nothing. |
| `--list-fields` | Print every text shape name and token in the template, then exit. Needs no data file. |
| `--overwrite` | Replace existing output files instead of adding `(2)`, `(3)`, and so on. |
| `-v`, `--verbose` | Print every field that is filled, and every text shape that is left alone. |
| `-q`, `--quiet` | Print errors only. |
| `--version` | Print the version and exit. |

Quote paths that contain spaces.

### What a run prints

Each step announces itself before it starts, so a slow run is never a blank screen. The PDF step is the one that usually takes seconds, and it says so:

```
$ python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"
Reading Deal data.yaml
Opening Template.pptx
Searching 67 text shapes for 3 IDs
Matched 3 IDs to 3 text shapes
Writing 3 values into 3 text shapes
Setting 2 document properties
Created PPTX: Output\Template - Deal data - 2026-11-26.pptx
Exporting PDF with PowerPoint (this can take a few seconds)
Created PDF:  Output\Template - Deal data - 2026-11-26.pdf (3.1s)
```

- Every line is flushed as it is printed, so the output streams into a terminal, a log file, or a CI console while the run is still going.
- `-v` adds the detail behind those counts: every field that is written with its slide, layout, or master, every shape left unchanged, and every document property with its value.
- `-q` prints errors only, which is what you want when a parent script already logs the job.

Exit codes:

| Code | Meaning |
| --- | --- |
| 0 | Success. |
| 1 | Unexpected failure. |
| 2 | Bad arguments, missing files, or invalid YAML/JSON. |
| 3 | The data file refers to IDs the template does not contain. |
| 4 | PDF export failed. The PPTX may still exist. |

## Check the template first

`--list-fields` prints every fillable text shape, where it sits, and its current text. It also shows which IDs are already written as `{{tokens}}`.

```
$ python ppt_template_to_pdf.py "Template.pptx" --list-fields
Where                  Shape name         Match by           Text
---------------------  -----------------  -----------------  --------------------
slide 1                CLIENT_NAME        shape name         Client name here
slide 1                TRANSACTION_TITLE  shape name         Project ___ Financing
slide 3                TextBox 7          token SUMMARY      {{SUMMARY}}
master 'Office Theme'  TextBox 4          token FOOTER_TEXT  {{FOOTER_TEXT}}
```

Use those names, or the IDs inside the braces, as top-level YAML or JSON keys.

Layout and master shapes appear in that list only when they hold a token, because a stock template carries dozens of placeholder prompts there. Those shapes are still searched and can still be filled by shape name, which is why the list ends with a count of what it hid:

```
63 text shape(s) on layouts and masters hold no tokens and are not listed. They are searched too and can still be filled by shape name.
```

Preview a real run without writing files:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml" --dry-run -v
```

## Prepare the template

### Method 1 (recommended): shape names

1. Open the template in PowerPoint.
2. Go to Home → Arrange → Selection Pane.
3. Double-click a shape name in the pane and rename it to the field ID.
4. Use the same ID as the top-level YAML or JSON key.

![Home => Arrange => Selection Pane...](assets/screenshots/home-arrange-selectionpane.png)

![PowerPoint with the Selection Pane open. The shape renamed to TRANSACTION_TITLE is selected, and the same name is highlighted in the pane.](assets/screenshots/selection-pane.png)

Names you can use:

```
CLIENT_NAME
TRANSACTION_TITLE
KEY_POINTS
```

The screenshot uses `assets/screenshots/demo-template.pptx`. Open that file, show the Selection Pane, and the four names appear as in the picture.

### Method 2: tokens in the text

1. Type a token such as `{{CLIENT_NAME}}` inside a text box.
2. Use the ID inside the braces as the top-level YAML or JSON key.
3. Leave the shape name alone. Tokens match on the text, not on the name.

![A text box on the slide containing the token {{DISCLAIMER}}. The shape is selected and the pane shows its default name TextBox 7.](assets/screenshots/token.png)

In the screenshot the shape is still called `TextBox 7`, and the script fills it because its text contains `{{DISCLAIMER}}`.

### Method 3: on a slide layout or the slide master

A field does not have to live on a slide. Put the ID, as a shape name or as a `{{token}}`, on a slide layout or on the slide master, and the value reaches every slide that inherits from it:

- A text box on the **slide master** appears on every slide of every layout, so the master is the place for a footer, a confidentiality line, or a company name that must show everywhere.
- A text box on a **slide layout** reaches every slide that uses that layout.

A slide placeholder only overrides the layout and master text when it has text of its own. Leave it empty and the filled value shows through. In PowerPoint terms, the same rules for inheritance apply whether the text sits on the slide, the layout, or the master, and the script treats those three places the same way.

### Matching rules

- Shape names match exactly first.
- Shape names then match case-insensitively, with a warning.
- Tokens match with or without inner spaces, and are case-insensitive. `{{ SUMMARY }}`, `{{SUMMARY}}`, and `{{Summary}}` all match the ID `SUMMARY`.
- Shapes inside groups are searched, at any depth.
- **All three layers are searched:** slides, slide layouts, and slide masters. A field on a slide, on the layout it inherits from, and on the master is three matches.
- **One ID can match several shapes, and every match is filled.** The same shape name on slides 2, 5, and 9, or the same token repeated in five boxes, is not a conflict: the value is written everywhere.
- If one shape is targeted by two IDs, the script warns and the later ID in the data file wins.
- A token match replaces the whole text frame, so `Prepared for {{ CLIENT_NAME }}` becomes just the value.

Tables, charts, and pictures have no text frame, so the script skips them. Put the ID on a real text box.

Templates can be `.pptx`, `.pptm`, `.potx`, `.ppt`, or `.pot`. The output is always `.pptx`, so macros are dropped.

Legacy `.ppt` and `.pot` files are converted to `.pptx` in a temporary folder with PowerPoint before the script fills them. The temporary copy is deleted at the end. The output name still uses the original template name, without the old extension.

## HTML inside YAML or JSON values

Write formatting in the data file instead of clicking through PowerPoint. Values may contain a small set of HTML tags; the script turns them into real PowerPoint runs, paragraphs and bullets. This works in JSON strings exactly as it does in YAML strings.

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
SLIDE1_TITLE: "<b>Bold text</b> and <u>underlined text</u>"
SLIDE1_Subtitle: "Text with no formatting"
COLOURED: '<font color="#C00000">Red text</font>'
SIZED: '<font size="14">14-point text</font>'
COMBINED: '<span style="color:#C00000; font-size:14pt"><b>Red, bold, 14-point text</b></span>'

Slide2_BULLETS: |
  <ul>
    <li>First bullet</li>
    <li>Second bullet with <b>bold text</b></li>
  </ul>

Slide3_multiline_steps: |
  <ol>
    <li>First step</li>
    <li>Second step</li>
  </ol>
```

To print a tag as text, escape the angle brackets:

```yaml
LITERAL_TAG: "&lt;b&gt;This is not bold&lt;/b&gt;"
```

Use single quotes around a YAML value that contains HTML with double-quoted attributes. Double-quoted YAML strings end at the first inner `"`. The alternative is a block value with `|`. In JSON, escape the inner double quotes as `\"`.

## Pure YAML format

The top level must be a mapping of `ID: content`. YAML and JSON use the same structure; only the punctuation differs. The one reserved key is `PROPERTIES`, which sets document metadata instead of a shape — see [Document properties](#document-properties).

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

Paragraph and run formatting that the data file does not set comes from the first paragraph and first run of the template text box.

Unknown keys are rejected rather than ignored. A typo like `colourr` reports `did you mean 'colour'?`.

## JSON data files

Pass a `.json` file anywhere a `.yaml` file goes. The extension picks the parser, and everything downstream is identical: the same IDs, the same shape-name and `{{token}}` matching, the same HTML tags, the same flags and exit codes.

```json
{
  "CLIENT_NAME": "Example Energy Pte. Ltd.",
  "TRANSACTION_TITLE": {
    "runs": [
      { "text": "Project " },
      { "text": "Sophia", "bold": true, "colour": "#C00000" },
      { "text": " Financing Proposal" }
    ]
  },
  "SUMMARY": "<b>THEBANK</b> is proposing a structured financing solution.",
  "DISCLAIMER": {
    "text": "Strictly private and confidential",
    "bold": true,
    "colour": "#C00000",
    "size": 9,
    "align": "center"
  }
}
```

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.json"
```

Why JSON: it is what APIs, web forms, and `json.dump()` already produce, and it needs no extra dependency.

Rules that JSON adds:

- Keys and strings use double quotes. Single quotes and bare keys are syntax errors.
- No comments. `//` and `/* */` are syntax errors, which the script points out at the exact line.
- No trailing comma after the last entry of an object or an array.
- Duplicate keys are rejected, like duplicate YAML keys.
- Escape a line break inside a string as `\n`; a raw line break is a syntax error.
- Numbers and booleans keep their type, so `"size": 9` and `"bold": true` work without quotes.
- Keys that YAML would read as booleans, such as `NO` or `ON`, are plain strings in JSON. `{"NO": "Norway"}` works with no quoting tricks.

Errors are reported in the same style as the YAML ones, with the line, the column, the offending line, and a fix:

```
ERROR: JSON SYNTAX ERROR
=================
File:    Deal data.json
Line:    2
Column:  27
Problem: Illegal trailing comma before end of object

  1 | {
> 2 |   "CLIENT_NAME": "Example",
                                ^
  3 | }

How to fix it:
JSON allows no comma after the last entry of an object or an array. Delete the comma before the closing brace or bracket.
```

### Converting between YAML and JSON

YAML to JSON:

```
python -c "import json, sys, yaml; json.dump(yaml.safe_load(open(sys.argv[1], encoding='utf-8')), open(sys.argv[2], 'w', encoding='utf-8'), indent=2, ensure_ascii=False)" "Deal data.yaml" "Deal data.json"
```

JSON to YAML:

```
python -c "import json, sys, yaml; yaml.safe_dump(json.load(open(sys.argv[1], encoding='utf-8')), open(sys.argv[2], 'w', encoding='utf-8'), allow_unicode=True, sort_keys=False)" "Deal data.json" "Deal data.yaml"
```

## Document properties

Shape fields are not the only target. A deck also carries **document properties**, and they are what a PDF shows under File > Properties > Description: Title, Author, Subject, Keywords, Created. This is why a PDF exported from PowerPoint opens with a document title such as `Q3 Results — Board Pack` instead of the file name.

Set them in the data file with a `PROPERTIES` block. It sits next to the shape IDs and is never matched against shapes:

```yaml
CLIENT_NAME: "Example Energy Pte. Ltd."

PROPERTIES:
  title: "Q3 Results — Board Pack"
  subject: "Quarterly review"
  author: "Corporate Finance"
  keywords: [q3, board pack, energy]
  category: "Client reporting"
  comments: "Generated from the deal tracker"
  last_modified_by: "Deal team"
  content_status: "Final"
  version: "1.2"
  revision: 3
  created: "2026-09-30T09:00:00"
```

The same block in JSON:

```json
{
  "CLIENT_NAME": "Example Energy Pte. Ltd.",
  "PROPERTIES": {
    "title": "Q3 Results — Board Pack",
    "author": "Corporate Finance",
    "keywords": ["q3", "board pack"]
  }
}
```

Run it exactly as before; the properties are written just before the PPTX is saved, so the PDF export already sees them:

```
python ppt_template_to_pdf.py "Template.pptx" "Deal data.yaml"
```

Supported properties:

| Property | Meaning | Value |
| --- | --- | --- |
| `title` | Document title. The one the PDF shows. | Text. |
| `subject` | Subject line. | Text. |
| `author` | Author or team. | Text. |
| `keywords` | Tags for search and catalogues. | Text, or a list joined with commas. |
| `comments` | Free notes. | Text. |
| `category` | Category, for example `Client reporting`. | Text. |
| `last_modified_by` | Name of the last editor. | Text. |
| `content_status` | Status, for example `Draft` or `Final`. | Text. |
| `identifier` | Reference or document number. | Text. |
| `language` | Language tag, for example `en-GB`. | Text. |
| `version` | Document version, kept as text. | Text or number. |
| `revision` | Revision count. | Whole number, `0` or more. |
| `created` | Creation timestamp. | ISO 8601 such as `2026-09-30T09:00:00`. |
| `modified` | Last-saved timestamp. | ISO 8601. |
| `last_printed` | Last-printed timestamp. | ISO 8601. |

Details that save time:

- Property names match case-insensitively and accept spaces, hyphens, or camelCase, so `Title`, `title`, and `TITLE` are the same key, and `Last Modified By`, `lastModifiedBy`, and `last_modified_by` are the same property.
- An unknown property is rejected with a suggestion (`did you mean 'title'?`) instead of being silently dropped.
- `keywords` may be a YAML/JSON list. It is joined with commas, because a PPTX stores keywords as one string.
- Dates take an ISO 8601 string, a YAML date, or a Python datetime. A timezone is converted to UTC and dropped, because PPTX stores plain timestamps.
- A data file that holds only a `PROPERTIES` block is valid. Nothing is filled, only the metadata is set.
- `PROPERTIES` is a reserved key. A shape renamed to `PROPERTIES` cannot be filled, so avoid that name.
- On export, desktop PowerPoint copies these values into the PDF `Info` dictionary and the XMP metadata, so `Title`, `Author`, `Subject`, and `Keywords` come out in the PDF as well.
- Use `-v` to see each property that is written.

## Use it from Python

The script is a normal importable module. Call `main()` with the same arguments you would pass on the command line:

```python
from ppt_template_to_pdf import main

exit_code = main(["Template.pptx", "Deal data.yaml", "--no-pdf", "-q"])
exit_code = main(["Template.pptx", "Deal data.json", "--no-pdf", "-q"])
```

Use that to fill a template inside an existing ETL job, a Flask view, a Django management command, or a notebook.

A loop over a spreadsheet makes one deck per row:

```python
import yaml
from pathlib import Path
from ppt_template_to_pdf import main

rows = [
    {"CLIENT_NAME": "Example Energy Pte. Ltd.", "TRANSACTION_TITLE": "Project Sophia"},
    {"CLIENT_NAME": "Northwind Ports Ltd.", "TRANSACTION_TITLE": "Project Harbour"},
]

for index, row in enumerate(rows, start=1):
    data_file = Path(f"deal-{index}.yaml")
    data_file.write_text(yaml.safe_dump(row, allow_unicode=True, sort_keys=False), encoding="utf-8")
    main(["Template.pptx", str(data_file), "--output-dir", "Output", "--no-pdf"])
```

The loop writes YAML. If your pipeline already produces JSON, replace `yaml.safe_dump(...)` with `json.dump(row, stream, ensure_ascii=False)` and name the file `.json`.

If you need lower-level control, `python-pptx` is the library underneath and everything here is plain Python: open the presentation, walk the shapes, set `text_frame.text`, save. This repo adds the data file layer (YAML or JSON), the token matching, the formatting rules and the PDF export on top of that.

## Troubleshooting

| Message or symptom | Cause | Fix |
| --- | --- | --- |
| `Template file not found` | Wrong path or missing file. | Check the path. Quote paths with spaces. |
| `Data file not found` | Wrong path. | Check the path and the `.yaml`, `.yml`, or `.json` extension. |
| `Template must be one of .pot, .potx, .ppt, .pptm, .pptx` | The file is not a PowerPoint file, or it is a `.pps` / `.ppsx` slide show. | Open the file in PowerPoint and save it as `.pptx`. |
| `Reading a .ppt or .pot template needs Windows...` | The template is the old binary format and PowerPoint is not available. | Run on Windows with PowerPoint installed, or convert the file once with PowerPoint (File → Save As → `.pptx`) or `soffice --convert-to pptx "template.ppt"`. |
| `Data file must be one of .json, .yaml, .yml` | The data file has another extension. | Rename the file or fix the extension. |
| `Output path is a file, not a folder` | `-o` points at a file. | Pass a folder path, or delete the file. |
| `--name must not contain ...` | The name holds a path separator or a character Windows forbids in a file name. | Use letters, digits, spaces, `-`, `_`, or `.`. Put folders in `--output-dir`. |
| `--name ... is a name Windows reserves` | The name is `CON`, `NUL`, `COM1`, `LPT1`, and so on. | Pick another name. |
| `--name must not start or end with a dot` | The name starts with `.` (a hidden file) or ends with `.` (Windows drops it). | Rename without the leading or trailing dot. |
| `YAML COMMENT ERROR` | A double-quoted value contains `colour="#C00000"`. YAML reads `#` as a comment and cuts the value there. | The message prints the fixed line. Wrap the value in single quotes or use a `\|` block. |
| `YAML SYNTAX ERROR` with line and column | Bad quotes, tabs, or indentation. | Read the quoted line and the suggestion under it. Use spaces, never tabs. |
| `found duplicate key` | The same ID appears twice. | Remove the duplicate or rename one of them. |
| `ID ... is not text` | A key such as `NO`, `ON`, or `YES` was read as a boolean, not as a word. | Quote the key: `'NO': content`. |
| `is empty` or `contains no IDs` | The data file has no entries. | Add at least one `ID: content` entry (or `"ID": content` in JSON). |
| `is not valid UTF-8 text` | The file was saved in a legacy code page. | Save the data file as UTF-8. A UTF-8 BOM is fine. |
| `the top level must be a mapping` | The data file is a list or a single value. | Start the file with `ID: value` lines at column zero, or one JSON `{ ... }` object. |
| `PROPERTIES must be a mapping` | The `PROPERTIES` key holds a string or a list instead of `property: value` lines. | Indent the properties under the key, or use a JSON object. |
| `unknown document property` | A property name is misspelled or not supported. | Use a suggested name. The message lists every supported property. |
| `must be an ISO 8601 date and time` | `created`, `modified`, or `last_printed` holds free text. | Use `2026-09-30T09:00:00`. |
| `must be a whole number` | `revision` is not a number. | Use `3`, not `third`. |
| `These YAML IDs were not found in the template` (or `These JSON IDs ...`) | Shape name differs from the data file key. The field is a table or a picture. | The message suggests close matches. Run `--list-fields` to copy the exact name. |
| `JSON SYNTAX ERROR` with line and column | A trailing comma, a single-quoted or bare key, a `//` comment, or a raw line break inside a string. | Read the quoted line and the suggestion under it. JSON needs double quotes, no comments, no trailing commas. |
| `JSON DUPLICATE KEY` | The same key appears twice in one object. | Remove the duplicate or rename one of them. |
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
| A field on a layout or the master changes the master, but one slide still shows the old text | That slide's placeholder holds its own text, so it overrides the layout and the master. | Delete the text inside the slide placeholder, or move the field onto the slide. |
| Text overflows the box | The template box is too small. | Enlarge the box or enable autofit in the template. |
| PDF fonts differ from the template | A template font is not installed on this machine. | Install the font, or save the template with an available font. |

The script never prints a traceback for expected problems. It prints `ERROR: ...` to stderr and returns a code from the table above. Check `$LASTEXITCODE` in PowerShell and `$?` in Bash.

## FAQ: update PowerPoint with Python

**How do I update PowerPoint text with Python?**
Install `python-pptx`, open the file with `Presentation("Template.pptx")`, walk `slide.shapes`, and set `shape.text_frame.text`. That is the low-level path. This repo wraps it: you keep the template, put the values in a YAML or JSON file, and run one command. See [What the script does](#what-the-script-does).

**Is there a script to update a PowerPoint template with variables?**
Yes — that is what `ppt_template_to_pdf.py` is. Variables are top-level YAML or JSON keys; the matching shape name or `{{token}}` in the deck is replaced with the value.

**Can I use JSON instead of YAML?**
Yes. Pass a `.json` file as the second argument; the extension selects the parser. The IDs, values, HTML formatting, flags and exit codes are identical — only the punctuation differs. See [JSON data files](#json-data-files). JSON needs no extra dependency, so it is a good fit when an API or a `json.dump()` job already produces the data.

**Can a PowerPoint template be updated dynamically, at runtime?**
Yes. The script is a single command with a meaningful exit code, so you can call it from Python, PowerShell, Bash, a cron job, Task Scheduler, GitHub Actions, Azure Pipelines, or Jenkins. Generate the data file — YAML or JSON — from your database or API, then fill the template.

**How do I replace `{{PLACEHOLDER}}` tokens in a PPTX file?**
Name the shapes with the placeholder text or type the tokens directly into the text boxes, then pass a YAML or JSON mapping of token → value. `--list-fields` shows which tokens the template already contains.

**How do I do a mail merge in PowerPoint?**
PowerPoint has no built-in mail merge for slides. The usual workaround is one data file per record plus this script in a loop — see [Use it from Python](#use-it-from-python).

**Can I generate PowerPoint slides from Excel or a database?**
Yes. Export each row to YAML or JSON (or build the structure in Python), then call the script once per row. The template's design, master slides and fonts are reused every time.

**Does it keep my template's formatting?**
Yes. The script writes text into existing boxes and takes unset paragraph and run formatting from the first paragraph and first run of that box, so the template's look stays. Explicit `bold`, `colour`, `size`, `bullet` and `align` values in the data file override it.

**Do I need Microsoft PowerPoint installed?**
Only for PDF export and for converting legacy `.ppt` / `.pot` templates. With `--no-pdf` and a `.pptx` template, `python-pptx` alone is enough — no Office, no Windows.

**Does it work on macOS or Linux?**
The PPTX step does. Add `--no-pdf`; PDF export uses the Windows COM interface of desktop PowerPoint. On Linux, convert the result with `soffice --headless --convert-to pdf "deck.pptx"`.

**Can it insert images, charts or tables?**
Not currently. The script fills text boxes and text frames, including grouped shapes. Tables and pictures have no text frame, so they are skipped. Use `python-pptx` directly for those, or a placeholder image with the same idea.

**Can I fill dozens of fields at once?**
Yes. One data file can hold every field for every slide. The IDs are matched across the whole presentation, including shapes inside groups.

**How do I set the PDF title and metadata?**
Add a `PROPERTIES` block to the data file with `title`, `subject`, `author`, and `keywords`. The values are written to the deck's document properties before the PDF export, so PowerPoint carries them into the PDF `Info` dictionary and XMP metadata. See [Document properties](#document-properties).

**Can I choose the output file name?**
Yes. `-n` / `--name "Q3 update"` writes `Q3 update.pptx` and `Q3 update.pdf` in the output folder. Without it the name is `<template> - <data> - <date>`, and `(2)`, `(3)` are added rather than overwriting an earlier file.

**Does it work if the field is defined on the slide master or on a layout?**
Yes. Slides, slide layouts, and slide masters are all searched. A field placed on the master is filled once and every slide that inherits it shows the value, which is how footers, confidentiality lines, and company names are handled. Placeholders on a slide only override the layout and master when they hold their own text, so leave them empty for inherited values.

**What happens if the same field appears more than once in the deck?**
Every occurrence is filled with the same value: the same shape name on several slides, the same token repeated in one text box or across slides, and a field that exists on both a slide and its layout. Nothing stops at the first match, and the verbose output lists each write.

**Why does the script seem to hang during the PDF step?**
Exporting with desktop PowerPoint takes a few seconds, and on a large deck it can take longer. The script prints `Exporting PDF with PowerPoint (this can take a few seconds)` before it starts, flushes it immediately, and reports the elapsed time when the PDF is written. If the line does not advance for a minute or more, close other Office windows and check that no dialog is waiting behind them — a hidden prompt blocks the export.

**Is there a way to check the template without running a real fill?**
`--list-fields` lists every text shape and token, and `--dry-run -v` reports the planned matches and writes nothing.

## How this compares to other tools

| Approach | What it gives you | Trade-off |
| --- | --- | --- |
| This repo (`ppt_template_to_pdf.py`) | Command line, YAML or JSON data file, `{{token}}` and shape-name matching, formatting, PPTX + PDF output, clear errors | Text fields only, no images or charts |
| `python-pptx` alone | Full control over shapes, tables, images, charts | You write and maintain the fill loop, the matching and the error handling |
| `python-pptx` + Jinja2 templates | Slide duplication and loops | You design the mapping and the formatting layer yourself |
| `docxtpl` | Great template filling — for Word `.docx` | Does not handle PowerPoint files |
| PowerPoint "Reuse Slides" / manual slides | No code | Not repeatable, not dynamic |
| Office Scripts / Power Automate | Cloud automation | Needs a Microsoft 365 licence and the cloud, and PowerPoint coverage is limited |

If your input is a spreadsheet and your output is a set of branded decks plus PDFs, this script is the short path.

## Search terms

Variables, placeholders, tokens, mail merge, templating, data merge, python-pptx wrapper, pptx template filler, PowerPoint automation, PowerPoint generation, PPTX to PDF, `.potx` template, `.pptm`, `python-pptx` text replacement, replace text in PowerPoint with Python, fill PowerPoint template from YAML, fill PowerPoint template from JSON, fill PPTX from an API response, PowerPoint template from JSON in Python, generate slides from JSON data, fill PowerPoint template from Excel, generate PowerPoint report in Python, dynamic PowerPoint decks, batch PowerPoint generation, PowerPoint CLI, slide master placeholders, edit slide layout text in Python, replace placeholder on every slide, fill every occurrence of a placeholder, set PowerPoint document properties in Python, set PDF title and author from Python, PDF metadata for PowerPoint export, python-pptx core properties, docProps title subject keywords.

## Pre-built executables

Every release carries a binary for each common platform, so a user with no Python installed can still run the fill:

| File | Runs on | Notes |
| --- | --- | --- |
| `ppt-fill-<version>-windows-x64.exe` | 64-bit Windows | The usual choice on Windows. |
| `ppt-fill-<version>-windows-x86.exe` | 32-bit Windows, and 64-bit Windows | Take it only if you are on 32-bit Windows or a tool of yours insists on a 32-bit binary. |
| `ppt-fill-<version>-macos-arm64` | macOS on Apple silicon (M1 and later) | Needs macOS 11 or newer. |
| `ppt-fill-<version>-macos-x86_64` | macOS on Intel | Needs macOS 11 or newer. |
| `ppt-fill-<version>-linux-x86_64` | 64-bit Linux | Needs glibc 2.35 or newer: Ubuntu 22.04+, Debian 12+, Fedora 36+, RHEL 10+. |

Each one is a single file: the CPython runtime, the standard library, python-pptx, PyYAML, lxml, Pillow, XlsxWriter, and (on Windows) pywin32 are all inside. Copy it anywhere, open a terminal, and run it like the script:

```
# Windows
ppt-fill-1.9.1-windows-x64.exe --help
ppt-fill-1.9.1-windows-x64.exe "Template.pptx" "Deal data.yaml"

# macOS and Linux: restore the executable bit first if the download lost it
chmod +x ./ppt-fill-1.9.1-linux-x86_64
./ppt-fill-1.9.1-linux-x86_64 "Template.pptx" "Deal data.yaml"
./ppt-fill-1.9.1-macos-arm64 "Template.pptx" --list-fields
./ppt-fill-1.9.1-linux-x86_64 "Template.pptx" "Deal data.json" -o Output --name "Q3 update"
```

Every flag, exit code, and message is the same as the Python version. Three things are worth knowing:

- **PDF export is Windows-only.** It drives desktop PowerPoint over COM, so on macOS and Linux the binary fills the `.pptx` and stops there. Add `--no-pdf` and convert afterwards with LibreOffice: `soffice --headless --convert-to pdf "deck.pptx"`. Legacy `.ppt` and `.pot` templates are Windows-only for the same reason.
- **One file, unpacked at start.** These are Nuitka one-file builds, so each one unpacks itself into the temp folder (`%TEMP%`, `$TMPDIR`, `/tmp`) before running. That costs a second or two on the first start and leaves nothing behind.
- **Downloaded binaries may need a nudge.** A file fetched through a browser can lose the executable bit on macOS and Linux (`chmod +x`), and macOS may quarantine it: `xattr -d com.apple.quarantine ppt-fill-1.9.1-macos-arm64`. The macOS builds are unsigned, so prefer `curl -LO` when you want to skip that.

Windows builds work with either bitness of Office, because PowerPoint is an out-of-process COM server: a 32-bit `ppt-template-to-pdf` can export a PDF through 64-bit Office and the other way round.

All of them are built in CI by [Nuitka](https://nuitka.net/) through [Nuitka-Action](https://github.com/Nuitka/Nuitka-Action), one job per platform and architecture (`ubuntu-22.04` for Linux, `macos-latest` and `macos-15-intel` for macOS, `windows-latest` for Windows). Each binary is smoke tested on its own runner before it is attached to the release: `--version`, `--list-fields`, and a real fill of the demo deck. Nuitka compiles the Python modules to C and bundles the runtime and the extension modules, which is what makes a single file possible.

To build one yourself, on the platform you are targeting:

```
python -m pip install -r requirements.txt nuitka
python -m nuitka --onefile --output-dir=dist --output-filename=ppt-template-to-pdf \
    --include-package-data=pptx ppt_template_to_pdf.py
```

On Windows add `--windows-icon-from-ico=assets/logo.ico --windows-console-mode=force --include-package=win32com` and finish the name with `.exe`. Nuitka cannot cross-compile, so each platform and architecture needs its own interpreter: a 32-bit Python for the Windows x86 build, an Intel Mac for the macOS Intel build.

## Releases

Each release publishes six assets:

| Asset | What it is |
| --- | --- |
| `ppt-template-to-pdf-<version>.zip` | The lean source bundle, described below. |
| `ppt-fill-<version>-windows-x64.exe` | Windows, 64-bit. |
| `ppt-fill-<version>-windows-x86.exe` | Windows, 32-bit. |
| `ppt-fill-<version>-macos-arm64` | macOS on Apple silicon. |
| `ppt-fill-<version>-macos-x86_64` | macOS on Intel. |
| `ppt-fill-<version>-linux-x86_64` | Linux, 64-bit. |

See [Pre-built executables](#pre-built-executables) for requirements and usage.

The zip holds only what a user needs, laid out inside one folder:

```
ppt-template-to-pdf-1.9.1/
  ppt_template_to_pdf.py
  README.md
  pyproject.toml
  requirements.txt
  LICENSE
```

Tests, the `assets/` folder (screenshots, logo, demo template), and `requirements-dev.txt` are left out on purpose, so the download stays small and nothing development-only reaches a user.

Build the same zip locally:

```
python tools/make_release_zip.py
```

It writes `dist/ppt-template-to-pdf-<version>.zip`, reads the version from `pyproject.toml`, and refuses to build when that version and `__version__` in the script disagree.

Cut a release by pushing a tag that matches the version:

```
git tag v1.9.1
git push origin main --tags
```

`.github/workflows/release.yml` then does the rest:

- **version** works out the tag and checks it looks like `1.9.1`.
- **bundle** runs the test suite, builds the zip, fails if anything unexpected slipped in, installs the zip in a clean runner and runs `ppt-template-to-pdf --version`.
- **windows** builds the executable twice with Nuitka, once per architecture (`x64`, `x86`), on `windows-latest`.
- **macos** builds it on `macos-latest` (Apple silicon) and `macos-15-intel`.
- **linux** builds it once on `ubuntu-22.04`, the oldest runner, which sets the glibc floor.
- Each build job smoke tests its own binary (`--version`, `--list-fields`, a real fill) before **publish** collects all six files and creates the GitHub release.

To release without pushing a tag by hand, run the workflow from the Actions tab with the version as the input; it creates the `v` tag for you. Tick `draft` to review the release before it goes out.

The workflow runs the test suite first, so a broken build never becomes a release. A tag can be released once: to redo one, delete its release and tag, then push the tag again.

## Run the tests

```
python -m pip install -r requirements-dev.txt
python -m pytest
```

The tests cover YAML and JSON loading, the HTML parser, matching, document properties, and PPTX writing, plus the release zip builder. They do not need PowerPoint. PDF export is checked by hand.

## License

MIT. See [LICENSE](LICENSE).
