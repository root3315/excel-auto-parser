# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Excel Smart Parser — universal Excel file parser that auto-detects and extracts all tables without configuration. Supports `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`. Includes a standalone web viewer (`excel_viewer/index.html`).

## Commands

```bash
# Install dependencies
pip install openpyxl xlrd pyxlsb chardet tqdm

# Run parser
python excel_smart_parser.py file.xlsx
python excel_smart_parser.py file.xlsx --format jsonl --stream --out-dir output/

# Run tests (150 unit tests)
python TEST/test_all_features.py
```

## Architecture

Single-file parser (`excel_smart_parser.py`, ~1900 lines) with adapter pattern:

**SheetAdapter (ABC)** — abstract interface for reading cells, dimensions, merged cells, hidden state:
- `OpenpyxlAdapter` — .xlsx, .xlsm
- `XlrdAdapter` — .xls (legacy)
- `PyxlsbAdapter` — .xlsb (binary)
- `CsvAdapter` — .csv (auto-detects encoding and delimiter)

**ExcelParser** — main orchestrator. For each sheet, extracts tables from 5 sources in priority order:
1. `native_table` — Excel Ctrl+T tables
2. `named_range` — named ranges
3. `heuristic` — score-based header detection (threshold configurable, default 0.4)
4. `vertical` — vertical tables (headers in column A)
5. `headerless` — raw data matrices

Uses a `used_rows` set to prevent overlapping extractions across sources.

**StreamingWriter** — writes JSONL/CSV incrementally for large files without loading all data into memory.

**Header detection** scores each candidate row (0.0–1.0) based on text type, length, date patterns, numeric ratio. Rows scoring above `header_threshold` become table headers.

## Key dependencies

- `openpyxl` (required) — Excel reading
- `xlrd` (optional) — .xls support
- `pyxlsb` (optional) — .xlsb support
- `chardet`/`charset-normalizer` (optional) — CSV encoding detection
- `tqdm` (optional) — progress bar
- `xlwt` — testing only (creating .xls fixtures)
