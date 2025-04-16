# PaperForge 
Markdown to DOCX ToolsKit

A collection of tools for converting Markdown documents into polished Word documents with advanced validation and testing capabilities.

## Overview

This project provides three main tools:

1. **Markdown to DOCX Generator** - Converts Markdown files to properly formatted Word documents
2. **DOCX Format Tester** - Examines and reports on the formatting details of a Word document
3. **Validation Pipeline** - Combines generation and testing to ensure proper conversion

These tools are useful for:
- Document automation
- Report generation
- Content management systems
- CV/Resume creation
- Academic paper preparation
- Documentation workflows

## Features

- **Rich formatting support**:
  - Headings and sections
  - Bold, italic text
  - Bulleted and numbered lists (including nested lists)
  - Tables
  - Code blocks with syntax highlighting
  - Links
  - Images
  - Horizontal rules

- **Validation capabilities**:
  - Style verification
  - Structural analysis
  - Format testing
  - Expected formatting rules validation

## Requirements

- Python 3.6+
- python-docx library (`pip install python-docx`)
- re (standard library)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/goliathuy/PaperForge.git  
   cd PaperForge
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Generate a Word document from Markdown

To generate a Word document from a Markdown file, use the following command:

```bash
python scripts/generate_word_from_md.py <input_md_file> <output_docx_file>
```

For example:

```bash
python scripts/generate_word_from_md.py examples/sample.md output.docx
```

### Test the formatting of a Word document

```bash
python scripts/test_docx_format.py <docx_file_path>
```

### Validate a Markdown to Word conversion

```bash
python scripts/validate_md_to_docx.py <input_md_file> [--analyze-only]
```

The `--analyze-only` flag will analyze the Markdown file and show expected formatting rules without generating a DOCX file.

## Example

```bash
# Generate a Word document
python scripts/generate_word_from_md.py examples/sample.md output.docx

# Test the formatting
python scripts/test_docx_format.py output.docx

# Validate the conversion
python scripts/validate_md_to_docx.py examples/sample.md
```

## How It Works

1. **generate_word_from_md.py**: Parses Markdown syntax and converts it to the appropriate Word formatting using python-docx. It handles various Markdown elements and creates a properly styled DOCX file.

2. **test_docx_format.py**: Analyzes a DOCX file and provides detailed information about paragraph styles, runs, formatting attributes, tables, and other elements.

3. **validate_md_to_docx.py**: Combines the generation and testing capabilities to perform a comprehensive validation, ensuring that the Markdown is correctly converted to the expected Word format.

## Customizing Styles

The Word documents are generated with a set of predefined styles that match common Markdown elements. If you need to customize these styles, you can modify the `generate_word_from_md.py` script.

## Limitations

- Complex nested structures might not convert perfectly
- Some advanced Markdown extensions may not be fully supported
- Image paths must be accessible from the script location
- Table formatting is basic

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT
