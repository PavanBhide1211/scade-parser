# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] — 2026-05-23

### Added

- Initial release of the SCADE Model Parser.
- `ScadeModelParser` class: XXE-safe recursive XML tree walk using `defusedxml`
  (with stdlib fallback that rejects DOCTYPE / external entities).
- Support for all three SCADE XML interface declaration patterns:
  - Pattern A — flat `Name`/`TypeRef` attributes on `<Input>` / `<Output>` elements.
  - Pattern B — `<Variable>` child elements inside `<Input>` / `<Output>`.
  - Pattern C — `<Variable Direction="in|out">` attribute style.
- Extraction of `<Sensor>` elements as calibration parameters (runtime-tunable).
- Extraction of `<Const>` and `<Constant>` elements as calibration parameters
  (compile-time constants). Both spellings supported explicitly.
- Default value resolution from element attributes (`Default`, `Value`, `InitVal`)
  and child text elements (`<Value>`, `<Default>`, `<InitVal>`).
- Multi-attribute fallback in `_attr()` for resilience across SCADE version differences
  (e.g. `TypeRef` vs `typeRef` vs `Type`).
- XML namespace stripping (`_strip_ns()`) for version-agnostic tag matching.
- `parse_path()` aggregator: handles single files and recursive directory scans.
  Per-file errors are caught and reported to `stderr`; processing continues.
- File-size pre-check: files > 50 MB are rejected before XML parsing begins.
- Extension allowlist: only `.scade` and `.xscade` files are opened.
- Output path sanitisation: filename restricted to `[A-Za-z0-9_.-]`; parent
  directory must pre-exist; `.xlsx` extension enforced.
- `write_excel()`: formatted multi-sheet `.xlsx` workbook (openpyxl).
  - **Summary** sheet: generation timestamp, output filename, record counts.
  - **Input Signals** sheet: Name, Type, Operator, Package, Comment, Source File.
  - **Output Signals** sheet: same columns as Input Signals.
  - **Calibration Parameters** sheet: Name, Type, Kind, Default Value, Operator,
    Package, Comment, Source File.
  - All data sheets: navy header row, alternating row fill, frozen header,
    auto-filter, auto-sized columns (capped at 55 chars).
- CLI via `argparse`: `INPUT` positional argument, `--output / -o` flag, `--help`.
- `sample_model.scade`: example model file covering all supported XML patterns.
- Full documentation suite in `docs/`:
  - Design Document (`.docx` + `.pdf`)
  - Architecture Document (`.docx` + `.pdf`)
  - User Guide (`.docx` + `.pdf`)
- GitHub Actions CI workflow: matrix build across Python 3.10 / 3.11 / 3.12
  on `ubuntu-latest` and `windows-latest`.

---

## [Unreleased]

_No unreleased changes yet._

---

<!-- Links -->
[1.0.0]: https://github.com/pavanbhide/scade-parser/releases/tag/v1.0.0
[Unreleased]: https://github.com/pavanbhide/scade-parser/compare/v1.0.0...HEAD
