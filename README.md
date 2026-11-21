# GEDCOM JSON → Word Document Generator

> **Mac users:**
> - Just want to use the app, nothing installed? → **[README_MAC.md](README_MAC.md)** (download a ready-made app)
> - Have Python and want to run from source? → **[README_MAC_PYTHON.md](README_MAC_PYTHON.md)**

Convert GEDCOM genealogy data (as JSON-LD) into formatted Word documents (.docx) with full person profiles, including birth/death events, relationships, and media file path references.

## Overview

This Python tool implements a **2-step pipeline**:

1. **Extract** (`extract_report`): Parse a JSON-LD file (e.g., from GEDCOM converter), filter `foaf:Person` objects from the `@graph`, and write an intermediate structured report with all person data and nested attributes properly hierarchized.

2. **Generate** (`generate_docx`): Read the extracted report and build a formatted .docx file with one section per person, including:
   - Person name, ID, genealogy type
   - All metadata (birth, death, events, relationships, occupations)
   - Nested attributes with proper indentation
   - Media/image file paths embedded at the location of image references

## Setup

### Prerequisites

- Python 3.8+
- pip

### Installation

1. Create a virtual environment (optional but recommended):

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

This installs:
- `python-docx` – Word document generation
- `PySide6` – Desktop GUI (Qt6)

## Usage

### Command-Line Interface (CLI)

Generate a report and DOCX in one command:

```bash
python python_poc/generator.py --json sources_files/Hrysio.json --report report/report.json --out target/new_output.docx
```

**Parameters:**
- `--json`: Path to input JSON-LD file (default: `python_poc/test.json`)
- `--report`: Path to save intermediate extracted report (default: `python_poc/extracted_report.json`)
- `--out`: Path to save final DOCX file (default: `python_poc/output.docx`)

**Example:**

```bash
python generator.py --json sample_family.json --report extracted.json --out family_tree.docx
```

### Graphical User Interface (GUI)

Run the PySide6 GUI to browse for files interactively:

```bash
python app.py
```

The GUI allows you to:
- Select a JSON input file via file browser
- Choose output DOCX location
- Click "Generate" to process
- View status and error messages

## Output

### Intermediate Report (`extracted_report.json`)

A structured JSON file containing:
- `sourceFile`: Path to the original JSON input
- `people`: Array of person objects with:
  - `order`: Sequence number in the document
  - `id`: GEDCOM ID (e.g., `_:731`)
  - `type`: GEDCOM type (e.g., `foaf:Person`)
  - `title`: Person's name
  - `lines`: Hierarchically formatted text lines of all attributes

Example entry (excerpt):

```json
{
  "order": 1,
  "id": "_:731",
  "type": "foaf:Person",
  "title": "Didier Michel/BEDER/",
  "lines": [
    "_FIL: LEGITIMATE_CHILD",
    "foaf:name:",
    "  GEDVALUE: Didier Michel/BEDER/",
    "  GIVN: Didier Michel",
    "  SURN: BEDER",
    "foaf:gender: M",
    "bio:event:",
    "  dc:date: 31 AUG 1962",
    "  @type: bio:Birth",
    "OBJE:",
    "  - Objet 1",
    "    @id: _:794",
    "    TITL: Birth Certificate",
    "    FILE: /path/to/birth_cert.jpg"
  ]
}
```

### Final DOCX (`output.docx`)

A formatted Word document with:
- **Header**: Title and page numbering (e.g., "GEDCOM JSON report")
- **Body**: One numbered section per person (in JSON graph order)
  - Person title in bold (14pt)
  - Metadata line in italics (ID and type)
  - Hierarchically indented attributes and events
  - Media file paths shown at the location of the image reference
  - Blank separator line between persons
- **Footer**: Auto-numbered pages

## Features

- **Ordered**: People appear in the same order as the `@graph` in the JSON input
- **Complete**: All attributes extracted, including nested objects and arrays
- **Hierarchical**: Proper indentation shows nesting relationships
- **Media-aware**: Image and document references shown as file paths (not embedded)
- **Robust encoding**: Automatically detects UTF-8, UTF-16, UTF-8-BOM, Latin-1
- **CLI + GUI**: Command-line for scripting, GUI for interactive use

## Example Workflow

### 1. Convert GEDCOM to JSON

If you have a GEDCOM file, first convert it to JSON-LD using the included NodeJS converter:

```bash
cd converterGedJson
node convert.js myFamilyTree.ged > ../python_poc/my_family.json
cd ../python_poc
```

### 2. Generate Report and Document

```bash
python generator.py --json my_family.json --report my_family_report.json --out my_family_tree.docx
```

### 3. Open in Word

Open `my_family_tree.docx` in Microsoft Word or compatible software.

## Technical Details

### JSON Input Format

Expected structure (from GEDCOM JSON-LD converters):

```json
{
  "@context": { "foaf": "...", "bio": "...", ... },
  "@graph": [
    {
      "@id": "_:ID",
      "@type": "foaf:Person",
      "foaf:name": { "GEDVALUE": "Name", "GIVN": "Given", ... },
      "bio:event": [ { "@type": "bio:Birth", "dc:date": "...", ... } ],
      "OBJE": [ { "@id": "_:OBJ", "FILE": { "GEDVALUE": "/path/to/file" } } ],
      ...
    },
    ...
  ]
}
```

### Processing Steps

1. **Load JSON** with auto-encoding detection
2. **Build node index**: Map `@id` to each node for fast lookups
3. **Extract people**: Filter `@graph` for `foaf:Person` entries
4. **Hierarchize**: Recursively format nested attributes with indentation
5. **Resolve references**: Look up media paths and linked objects by ID
6. **Generate DOCX**: Build document with person sections and formatting

## Troubleshooting

### `UnicodeDecodeError` on JSON load

The generator automatically tries multiple encodings (UTF-8, UTF-16, UTF-8-BOM, Latin-1). If it still fails, manually convert the JSON file to UTF-8:

```bash
# On Windows (PowerShell)
(Get-Content input.json -Encoding UTF16) | Set-Content input_utf8.json -Encoding UTF8

# On macOS/Linux
iconv -f UTF-16 -t UTF-8 input.json -o input_utf8.json
```

### Missing file paths in the DOCX

Ensure the input JSON contains `OBJE` (media object) entries with `FILE` and `GEDVALUE` fields. Image paths are extracted from the JSON, not from your file system (images are not embedded).

### Empty or no person sections

Verify the JSON has a `@graph` array with at least one `foaf:Person` entry (check `@type` field).

## Dependencies

- **python-docx** (0.8.11+): Creates and formats .docx files
- **PySide6** (6.5+): Qt6-based GUI (optional for CLI-only use)

## Files

- `generator.py` – Main CLI and extraction/generation logic
- `app.py` – PySide6 GUI application
- `requirements.txt` – Python package dependencies
- `README.md` – This file
- `test.json` – Sample GEDCOM JSON-LD input (provided)
- `extracted_report.json` – Generated intermediate report
- `output.docx` – Generated Word document

## License

(Add license information if applicable)

## Notes

- The generator preserves character encoding from the source JSON
- Image file paths are shown as text; images are not embedded in the DOCX (as requested)
- Nested structures can be quite deep; indentation goes up to 14pt per level
- The intermediate report file can be re-used to generate multiple DOCX files with different formatting (future enhancement)
