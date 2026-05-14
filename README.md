# slide-generator

A Python tool that takes a PowerPoint slide template containing empty tables and fills them with your data to produce a presentation-ready `.pptx` deck.

## Features

- Fill table cells from a simple JSON data file
- Two fill modes: **flat key→value placeholders** and **table-indexed row data**
- Preserves all template formatting (fonts, colors, sizes)
- CLI and Python API
- No PowerPoint installation required (`python-pptx` only)

## Installation

```bash
pip install -e ".[dev]"
```

## Quick start

### 1. Create a template

Use any PowerPoint app to create a `.pptx` with tables.  
Put `{{placeholder_name}}` tokens in cells you want filled dynamically.

Or generate the bundled sample template:

```bash
python examples/create_sample_template.py
```

### 2. Create a data file

```json
{
  "company": "Acme Corp",
  "quarter": "Q1 2026",
  "revenue": "$4.2M",

  "slide2_table1": [
    ["Region", "Revenue", "Growth"],
    ["North",  "$1.4M",   "+22%"],
    ["South",  "$0.9M",   "+11%"]
  ]
}
```

### 3. Generate the deck

**CLI:**
```bash
slide-generator templates/quarterly_report_template.pptx examples/quarterly_report.json
# → output/quarterly_report_template.out.pptx
```

**Python API:**
```python
from slide_generator import SlideGenerator

gen = SlideGenerator("templates/quarterly_report_template.pptx")
gen.generate(data, "output/my_deck.pptx")
```

## Data format

### Flat placeholders

Any `{{KEY}}` token found in a table cell or text frame is replaced with the corresponding value:

```json
{ "revenue": "$4.2M", "growth": "+18%" }
```

### Table-indexed data

Target a specific table using the key `slide<N>_table<M>` (both 1-based).  
Supply a 2-D list of rows; if the template table has a header row it is preserved automatically:

```json
{
  "slide2_table1": [
    ["North", "$1.4M", "+22%"],
    ["South", "$0.9M", "+11%"]
  ]
}
```

## Running tests

```bash
pytest
```

## Project structure

```
slide-generator/
├── slide_generator/
│   ├── __init__.py        # Public API
│   ├── generator.py       # SlideGenerator class
│   ├── table_filler.py    # Table fill logic
│   └── cli.py             # CLI entry point
├── templates/             # Place your .pptx templates here
├── examples/              # Sample data + template creation script
├── output/                # Generated decks (git-ignored)
├── tests/
│   └── test_generator.py
├── pyproject.toml
└── LICENSE                # MIT
```

## License

MIT — see [LICENSE](LICENSE).
