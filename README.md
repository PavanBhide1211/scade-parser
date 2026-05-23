# SCADE Model Parser

[![CI](https://github.com/pavanbhide/scade-parser/actions/workflows/ci.yml/badge.svg)](https://github.com/pavanbhide/scade-parser/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A command-line Python tool that parses **SCADE Suite** model files (`.scade` / `.xscade`) and extracts **input signals**, **output signals**, and **calibration parameters** (Sensor / Const elements) into a formatted, multi-sheet **Excel workbook**.

---

## Features

- Extracts input signals, output signals, and calibration parameters (Sensor, Const, Constant)
- Handles all three SCADE XML interface declaration patterns (flat attribute, Variable children, Direction attribute)
- Recursively scans directories for `.scade` and `.xscade` files
- Outputs a formatted `.xlsx` workbook with four sheets: Summary, Input Signals, Output Signals, Calibration Parameters
- **XXE-safe** XML parsing via [`defusedxml`](https://github.com/tiran/defusedxml) with a stdlib fallback
- Rejects files exceeding 50 MB (configurable) to guard against XML bomb attacks
- Extension allowlist — only opens `.scade` / `.xscade` files
- Single-file, no build step required — just copy and run

---

## Requirements

| Dependency   | Version   | Notes                                      |
|--------------|-----------|--------------------------------------------|
| Python       | ≥ 3.10    |                                            |
| openpyxl     | ≥ 3.1.0   | Required — Excel output                    |
| defusedxml   | ≥ 0.7.1   | **Strongly recommended** — XXE protection  |

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/pavanbhide/scade-parser.git
cd scade-parser

# 2. Install dependencies
pip install -r requirements.txt
```

No build step, no package installation needed — `scade_parser.py` is self-contained.

---

## Quick Start

```bash
# Parse a single file (output name auto-generated)
python scade_parser.py my_model.scade

# Parse an entire directory of models
python scade_parser.py ./models/

# Specify the output file
python scade_parser.py ./models/ --output ECU_Signals.xlsx
```

### Sample output

```
SCADE Model Parser
==================================================
Input  : ./models/
Output : ./ECU_Signals.xlsx

  Parsing : BrakeController.scade
  Parsing : ThrottleController.scade
  Parsing : Diagnostics.scade

Results
------------------------------
  Input signals          : 12
  Output signals         :  8
  Calibration parameters : 15

Writing Excel workbook …
  Sheet 'Input Signals': 12 rows
  Sheet 'Output Signals': 8 rows
  Sheet 'Calibration Parameters': 15 rows

✓  Saved: ./ECU_Signals.xlsx
```

---

## CLI Reference

```
usage: scade_parser [-h] [--output OUTPUT] INPUT

positional arguments:
  INPUT                 Path to a .scade file or directory containing .scade files.

options:
  -h, --help            Show this help message and exit.
  --output, -o OUTPUT   Output .xlsx file path.
                        Default: scade_signals_YYYYMMDD_HHMMSS.xlsx
```

### Exit codes

| Code | Meaning |
|------|---------|
| `0`  | Success — Excel file written. Also returned when all files were skipped or no signals were found (check stderr for `[SKIP]` / `[WARNING]` messages and verify the output file was created). |
| `1`  | Fatal error — invalid arguments, unwritable output path, or Excel write failure. |

> **CI tip:** Even on exit code 0, check that the output `.xlsx` was actually created, since all files could have been skipped due to per-file errors reported on stderr.

---

## Output Workbook

The generated Excel file contains four sheets:

| Sheet | Contents |
|-------|----------|
| **Summary** | Generation timestamp, output filename, and record counts per category. |
| **Input Signals** | Name, Type, Operator, Package, Comment, Source File |
| **Output Signals** | Name, Type, Operator, Package, Comment, Source File |
| **Calibration Parameters** | Name, Type, Kind (Sensor/Constant), Default Value, Operator, Package, Comment, Source File |

All data sheets include a frozen header row, alternating row fill, auto-filter, and auto-sized columns.

---

## Supported SCADE XML Patterns

The parser handles the following XML structures automatically:

```xml
<!-- Pattern A: flat attribute -->
<Input  Name="BrakePedalPos" TypeRef="float32"/>
<Output Name="BrakePressure" TypeRef="float32"/>

<!-- Pattern B: Variable children -->
<Input>
  <Variable Name="EngineSpeed" TypeRef="int32"/>
</Input>

<!-- Pattern C: Direction attribute -->
<Variable Name="ThrottleCmd" TypeRef="float32" Direction="out"/>

<!-- Calibration: Sensor (runtime-tunable) -->
<Sensor Name="CAL_MaxSpeed" TypeRef="float32">
  <Default>120.0</Default>
</Sensor>

<!-- Calibration: Const / Constant (compile-time) -->
<Const    Name="K_GAIN" TypeRef="float32" Value="2.5"/>
<Constant Name="K_PI"   TypeRef="float64"><Value>3.14159</Value></Constant>
```

---

## Security

This tool is designed to process model files safely in automated pipelines.

| Threat | Mitigation |
|--------|-----------|
| XML External Entity (XXE) injection | `defusedxml` blocks all DOCTYPE / entity declarations. Stdlib fallback rejects external entity resolvers — **partial protection only; install `defusedxml` in all non-local environments**. |
| XML bomb (billion-laughs) | `defusedxml` detects recursive entity expansion. File-size pre-check (50 MB cap) provides additional defence. |
| Path traversal via `--output` | Output path parent directory must pre-exist; filename is sanitised to `[A-Za-z0-9_.-]` only. |
| Oversized / corrupted files | Files > 50 MB are rejected before parsing begins. |
| Dynamic code execution | No `eval()`, `exec()`, or `shell=True` anywhere in the codebase. |

See [SECURITY.md](SECURITY.md) for the vulnerability reporting process.

---

## Programmatic Use

`parse_path` and `write_excel` can be imported directly without going through the CLI:

```python
from scade_parser import parse_path, write_excel
from pathlib import Path

inputs, outputs, calibrations = parse_path(Path("./models/"))

# Post-process as needed, then write
write_excel(inputs, outputs, calibrations, Path("signals.xlsx"))

# Or work with the raw dicts directly
for signal in inputs:
    print(signal["Name"], signal["Type"], signal["Operator"])
```

---

## Documentation

Full documentation is available in the [`docs/`](docs/) folder:

| Document | Description |
|----------|-------------|
| [Design Document](docs/SCADE_Parser_Design.docx) | Requirements, data model, class design, error handling, design decisions |
| [Architecture Document](docs/SCADE_Parser_Architecture.docx) | Component layers, XML parsing strategy, security architecture, extension points |
| [User Guide](docs/SCADE_Parser_UserGuide.docx) | Installation, CLI reference, output interpretation, troubleshooting, FAQ |

PDF versions of each document are also included.

---

## Project Structure

```
scade-parser/
├── scade_parser.py          # Main parser (single-file, self-contained)
├── sample_model.scade       # Example SCADE XML file for testing
├── requirements.txt         # Python dependencies
├── pyproject.toml           # Package metadata
├── README.md
├── LICENSE                  # MIT
├── CHANGELOG.md
├── CONTRIBUTING.md
├── SECURITY.md
└── docs/
    ├── SCADE_Parser_Design.docx / .pdf
    ├── SCADE_Parser_Architecture.docx / .pdf
    └── SCADE_Parser_UserGuide.docx / .pdf
```

---

## Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a pull request.

---

## License

MIT © 2026 Pavan Bhide — see [LICENSE](LICENSE) for full terms.
