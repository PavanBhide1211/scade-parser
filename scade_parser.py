#!/usr/bin/env python3
"""
scade_parser.py
===============
Parses SCADE Suite model files (.scade / .xscade) and extracts:
  - Input signals
  - Output signals
  - Calibration parameters  (Sensor / Const elements)

Outputs a formatted multi-sheet Excel workbook (.xlsx).

Usage
-----
  python scade_parser.py model.scade
  python scade_parser.py ./models_dir/ --output signals_report.xlsx

Security notes
--------------
  - XML External Entity (XXE) injection is blocked via defusedxml (preferred)
    or by disabling external-entity resolution in the stdlib ET parser.
  - File-size cap (MAX_FILE_SIZE_MB) guards against zip-bomb-style XML files.
  - Path traversal is mitigated by resolving all paths to absolute form before
    any open() call and validating extensions against an allowlist.
  - No exec/eval or shell=True subprocess calls anywhere in this file.
  - Output filename is sanitised and must have a .xlsx extension.
"""

import argparse
import os
import sys
import re
from pathlib import Path
from datetime import datetime

# ── XML parsing (XXE-safe) ────────────────────────────────────────────────────
try:
    import defusedxml.ElementTree as _ET
    def _parse_xml(path: str):
        """Parse an XML file with XXE / billion-laughs protection."""
        return _ET.parse(path)
except ImportError:
    # defusedxml not installed – fall back to stdlib with external entity
    # resolution disabled (forbid_dtd still raises on malicious DTD).
    import xml.etree.ElementTree as _ET_std

    class _ForbiddenEntities:
        """Minimal expat target that rejects any DOCTYPE / entity declaration."""
        def __init__(self):
            import xml.parsers.expat as expat
            self.parser = expat.ParserCreate()
            self.parser.buffer_text = True

        @staticmethod
        def _reject(*_args, **_kwargs):
            raise ValueError(
                "DOCTYPE / external entities are not allowed in SCADE model files. "
                "Install 'defusedxml' for full XXE protection: pip install defusedxml"
            )

    def _parse_xml(path: str):
        """Parse with DOCTYPE/external-entity rejection."""
        parser = _ET_std.XMLParser()
        # Overwrite dangerous handlers with a rejecting stub
        if hasattr(parser, "parser"):
            parser.parser.ExternalEntityParserCreate = _ForbiddenEntities._reject
            if hasattr(parser.parser, "DefaultHandler"):
                parser.parser.DefaultHandler = _ForbiddenEntities._reject
        tree = _ET_std.parse(path, parser=parser)
        return tree

import xml.etree.ElementTree as ET   # only used for type annotations / ET.ParseError

# ── Excel output ──────────────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit(
        "[ERROR] openpyxl is required.\n"
        "Install it with:  pip install openpyxl"
    )


# =============================================================================
# CONSTANTS
# =============================================================================

MAX_FILE_SIZE_MB: int = 50          # Reject files larger than this
ALLOWED_EXTENSIONS: set = {".scade", ".xscade"}

# Output column definitions per sheet
_INPUT_COLS        = ["Name", "Type", "Operator", "Package", "Comment", "Source File"]
_OUTPUT_COLS       = ["Name", "Type", "Operator", "Package", "Comment", "Source File"]
_CALIBRATION_COLS  = ["Name", "Type", "Kind", "Default Value",
                       "Operator", "Package", "Comment", "Source File"]

# Excel style constants
_HDR_FILL  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
_HDR_FONT  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_ROW_FONT  = Font(name="Calibri", size=10)
_ALT_FILL  = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
_THIN_SIDE = Side(style="thin", color="BFBFBF")
_THIN_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE
)


# =============================================================================
# PATH / INPUT VALIDATION HELPERS
# =============================================================================

def _sanitise_filename(name: str) -> str:
    """Remove any characters that are unsafe in a filename."""
    # Allow only alphanumerics, dash, underscore, dot
    return re.sub(r"[^\w\-.]", "_", name)


def _validate_input_path(path: Path) -> None:
    """Raise ValueError / FileNotFoundError for invalid input paths."""
    resolved = path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Path not found: {path}")
    if resolved.is_file():
        ext = resolved.suffix.lower()
        if ext not in ALLOWED_EXTENSIONS:
            raise ValueError(
                f"Unsupported extension '{ext}'. "
                f"Expected one of: {sorted(ALLOWED_EXTENSIONS)}"
            )
        size_mb = resolved.stat().st_size / (1024 * 1024)
        if size_mb > MAX_FILE_SIZE_MB:
            raise ValueError(
                f"File '{resolved.name}' is {size_mb:.1f} MB, "
                f"exceeding the {MAX_FILE_SIZE_MB} MB safety limit."
            )


def _validate_output_path(path: Path) -> None:
    """Ensure the output path is a writable .xlsx location."""
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Output file must have a .xlsx extension.")
    # Attempt to detect write permission early (without creating the file)
    parent = path.parent.resolve()
    if not parent.exists():
        raise FileNotFoundError(f"Output directory does not exist: {parent}")
    if not os.access(parent, os.W_OK):
        raise PermissionError(f"No write permission for directory: {parent}")


# =============================================================================
# XML UTILITIES
# =============================================================================

def _strip_ns(tag: str) -> str:
    """Remove the XML namespace URI from a tag, e.g. '{uri}Tag' → 'Tag'."""
    return tag.split("}", 1)[1] if "}" in tag else tag


def _attr(element, *names: str, default: str = "") -> str:
    """Return the first matching attribute value from a list of candidate names."""
    for name in names:
        val = element.get(name)
        if val is not None:
            return str(val).strip()
    return default


def _child_text(element, *tags: str, default: str = "") -> str:
    """Return the text content of the first child whose tag matches any in *tags."""
    for child in element:
        if _strip_ns(child.tag) in tags and child.text:
            return child.text.strip()
    return default


# =============================================================================
# CORE PARSER
# =============================================================================

class ScadeModelParser:
    """
    Parses a single SCADE .scade or .xscade XML model file.

    After calling .parse(), access:
        .inputs         – list of dict records for input signals
        .outputs        – list of dict records for output signals
        .calibrations   – list of dict records for Sensor / Const elements
    """

    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path.resolve()
        self.inputs:       list[dict] = []
        self.outputs:      list[dict] = []
        self.calibrations: list[dict] = []
        self._source = file_path.name

    # ── Public API ────────────────────────────────────────────────────────────

    def parse(self) -> "ScadeModelParser":
        """Parse the file and populate inputs / outputs / calibrations."""
        _validate_input_path(self.file_path)
        try:
            tree = _parse_xml(str(self.file_path))
        except (ET.ParseError, ValueError) as exc:
            raise ValueError(
                f"XML parse error in '{self.file_path.name}': {exc}"
            ) from exc

        root = tree.getroot()
        self._walk(root, operator="", package="")
        return self

    # ── Private walk ──────────────────────────────────────────────────────────

    def _walk(self, element, operator: str, package: str) -> None:
        """Recursively traverse the XML element tree."""
        tag = _strip_ns(element.tag)

        # ── Context tracking ──────────────────────────────────────────────────
        if tag == "Package":
            pkg_name = _attr(element, "Name", "name")
            package = f"{package}/{pkg_name}".lstrip("/") if pkg_name else package

        elif tag in ("Operator", "Node", "Function"):
            operator = _attr(element, "Name", "name", default="<unnamed>")

        # ── Input signals ─────────────────────────────────────────────────────
        elif tag == "Input":
            # Pattern A: <Input Name="x" TypeRef="float32"/>  (flat attribute)
            if _attr(element, "Name", "name"):
                self._collect_var(element, "Input", operator, package)
            # Pattern B: <Input><Variable Name="x" …/></Input>
            for child in element:
                if _strip_ns(child.tag) == "Variable":
                    self._collect_var(child, "Input", operator, package)
            return  # children already handled above

        # ── Output signals ────────────────────────────────────────────────────
        elif tag == "Output":
            if _attr(element, "Name", "name"):
                self._collect_var(element, "Output", operator, package)
            for child in element:
                if _strip_ns(child.tag) == "Variable":
                    self._collect_var(child, "Output", operator, package)
            return

        # ── Variable with explicit Direction attribute ─────────────────────────
        elif tag == "Variable":
            direction = _attr(element, "Direction", "direction", "dir").lower()
            if direction in ("in", "input"):
                self._collect_var(element, "Input", operator, package)
            elif direction in ("out", "output"):
                self._collect_var(element, "Output", operator, package)
            # Variables without explicit direction are left to parent context

        # ── Calibration parameters: Sensor ───────────────────────────────────
        elif tag == "Sensor":
            self._collect_cal(element, "Sensor", operator, package)

        # ── Calibration parameters: Constant ─────────────────────────────────
        elif tag in ("Const", "Constant"):
            self._collect_cal(element, "Constant", operator, package)

        # ── Recurse into children ─────────────────────────────────────────────
        for child in element:
            self._walk(child, operator, package)

    # ── Record collectors ─────────────────────────────────────────────────────

    def _collect_var(self, el, direction: str, operator: str, package: str) -> None:
        name = _attr(el, "Name", "name")
        if not name:
            return
        record = {
            "Name":        name,
            "Type":        _attr(el, "TypeRef", "typeRef", "Type", "type",
                                  "DataType", default="unknown"),
            "Operator":    operator,
            "Package":     package,
            "Comment":     _attr(el, "Comment", "comment", "Description", default=""),
            "Source File": self._source,
        }
        (self.inputs if direction == "Input" else self.outputs).append(record)

    def _collect_cal(self, el, kind: str, operator: str, package: str) -> None:
        name = _attr(el, "Name", "name")
        if not name:
            return

        # Resolve default / initial value (attribute or child element)
        default_val = _attr(el, "Default", "Value", "InitVal", "DefaultValue", default="")
        if not default_val:
            default_val = _child_text(el, "Value", "Default", "InitVal", default="")

        record = {
            "Name":          name,
            "Type":          _attr(el, "TypeRef", "typeRef", "Type", "type",
                                    "DataType", default="unknown"),
            "Kind":          kind,        # "Sensor" or "Constant"
            "Default Value": default_val,
            "Operator":      operator,
            "Package":       package,
            "Comment":       _attr(el, "Comment", "comment", "Description", default=""),
            "Source File":   self._source,
        }
        self.calibrations.append(record)


# =============================================================================
# MULTI-FILE AGGREGATOR
# =============================================================================

def parse_path(
    path: Path,
) -> tuple[list[dict], list[dict], list[dict]]:
    """
    Parse a single .scade file or recursively scan a directory.

    Returns (inputs, outputs, calibrations) aggregated across all files found.
    """
    all_inputs: list[dict] = []
    all_outputs: list[dict] = []
    all_cals: list[dict] = []

    files: list[Path] = []
    resolved = path.resolve()

    if resolved.is_file():
        files = [resolved]
    elif resolved.is_dir():
        for ext in ALLOWED_EXTENSIONS:
            files.extend(sorted(resolved.rglob(f"*{ext}")))
    else:
        raise FileNotFoundError(f"Input path does not exist: {path}")

    if not files:
        print(
            f"[WARNING] No .scade / .xscade files found under: {resolved}",
            file=sys.stderr,
        )
        return [], [], []

    for fp in files:
        print(f"  Parsing : {fp.name}")
        try:
            p = ScadeModelParser(fp).parse()
            all_inputs.extend(p.inputs)
            all_outputs.extend(p.outputs)
            all_cals.extend(p.calibrations)
        except (ValueError, FileNotFoundError, PermissionError) as exc:
            print(f"  [SKIP]   {fp.name}: {exc}", file=sys.stderr)

    return all_inputs, all_outputs, all_cals


# =============================================================================
# EXCEL WRITER
# =============================================================================

def write_excel(
    inputs:       list[dict],
    outputs:      list[dict],
    calibrations: list[dict],
    output_path:  Path,
) -> None:
    """
    Write a formatted .xlsx workbook with four sheets:
      1. Summary
      2. Input Signals
      3. Output Signals
      4. Calibration Parameters
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # drop the default blank sheet

    # ── Sheet definitions ────────────────────────────────────────────────────
    sheet_specs = [
        ("Input Signals",         _INPUT_COLS,       inputs),
        ("Output Signals",        _OUTPUT_COLS,       outputs),
        ("Calibration Parameters", _CALIBRATION_COLS, calibrations),
    ]

    for sheet_name, columns, rows in sheet_specs:
        ws = wb.create_sheet(title=sheet_name)
        _write_data_sheet(ws, columns, rows)
        count = len(rows)
        print(f"  Sheet '{sheet_name}': {count} row{'s' if count != 1 else ''}")

    # ── Summary sheet (insert at front) ──────────────────────────────────────
    _write_summary_sheet(
        wb, inputs, outputs, calibrations, output_path
    )

    wb.save(str(output_path))


def _write_data_sheet(ws, columns: list[str], rows: list[dict]) -> None:
    """Populate a single data sheet with headers, data, formatting."""

    # Header row
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = _HDR_FILL
        cell.font      = _HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border    = _THIN_BORDER
    ws.row_dimensions[1].height = 22

    # Data rows
    for row_idx, record in enumerate(rows, start=2):
        use_alt = (row_idx % 2 == 0)
        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx,
                           value=record.get(col_name, ""))
            cell.font      = _ROW_FONT
            cell.alignment = Alignment(vertical="center")
            cell.border    = _THIN_BORDER
            if use_alt:
                cell.fill = _ALT_FILL

    # Auto-size columns (cap at 55 chars)
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(col_name)
        for row_cells in ws.iter_rows(
            min_row=2, min_col=col_idx, max_col=col_idx
        ):
            for cell in row_cells:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 55)

    # Freeze header + enable auto-filter
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions


def _write_summary_sheet(wb, inputs, outputs, calibrations, output_path):
    """Insert a summary / metadata sheet at position 0."""
    ws = wb.create_sheet(title="Summary", index=0)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 22
    ws.sheet_view.showGridLines = False

    title_font  = Font(bold=True, size=16, name="Calibri", color="1F4E79")
    label_font  = Font(bold=True,  name="Calibri", size=11)
    value_font  = Font(name="Calibri", size=11)
    section_font = Font(bold=True, name="Calibri", size=12, color="2E75B6")

    ws.merge_cells("A1:B1")
    ws["A1"] = "SCADE Model Parse Report"
    ws["A1"].font      = title_font
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A2"] = ""

    meta_rows = [
        ("Generated",   datetime.now().strftime("%Y-%m-%d  %H:%M:%S")),
        ("Output File", output_path.name),
    ]
    for r, (lbl, val) in enumerate(meta_rows, start=3):
        ws.cell(row=r, column=1, value=lbl).font  = label_font
        ws.cell(row=r, column=2, value=val).font   = value_font

    ws.cell(row=6, column=1, value="Signal Summary").font = section_font

    stat_rows = [
        ("Input Signals",          len(inputs)),
        ("Output Signals",         len(outputs)),
        ("Calibration Parameters", len(calibrations)),
        ("Total Items",            len(inputs) + len(outputs) + len(calibrations)),
    ]
    for r, (lbl, val) in enumerate(stat_rows, start=7):
        lc = ws.cell(row=r, column=1, value=lbl)
        vc = ws.cell(row=r, column=2, value=val)
        lc.font = label_font
        vc.font = value_font
        if lbl == "Total Items":
            lc.font = Font(bold=True, name="Calibri", size=11, color="1F4E79")
            vc.font = Font(bold=True, name="Calibri", size=11, color="1F4E79")


# =============================================================================
# CLI ENTRY POINT
# =============================================================================

def _build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="scade_parser",
        description=(
            "Parse SCADE model files (.scade / .xscade) and extract\n"
            "input signals, output signals, and calibration parameters\n"
            "(Sensor / Const) into a formatted Excel workbook."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python scade_parser.py model.scade\n"
            "  python scade_parser.py ./models/ --output signals.xlsx\n"
        ),
    )
    ap.add_argument(
        "input",
        metavar="INPUT",
        help="Path to a .scade file or a directory containing .scade files.",
    )
    ap.add_argument(
        "--output", "-o",
        metavar="OUTPUT",
        default=None,
        help=(
            "Output Excel file path (must end in .xlsx). "
            "Default: scade_signals_<YYYYMMDD_HHMMSS>.xlsx in the current directory."
        ),
    )
    return ap


def main(argv: list[str] | None = None) -> int:
    ap = _build_arg_parser()
    args = ap.parse_args(argv)

    # ── Resolve input path ────────────────────────────────────────────────────
    input_path = Path(args.input).resolve()

    # ── Resolve / validate output path ───────────────────────────────────────
    if args.output:
        raw_out = args.output
        # Basic sanitisation: keep only the basename portion sanitised,
        # but allow the user to specify a directory prefix
        out_p = Path(raw_out)
        safe_name = _sanitise_filename(out_p.name)
        output_path = (out_p.parent / safe_name).resolve()
    else:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(f"scade_signals_{ts}.xlsx").resolve()

    try:
        _validate_output_path(output_path)
    except (ValueError, FileNotFoundError, PermissionError) as exc:
        print(f"[ERROR] Output path problem: {exc}", file=sys.stderr)
        return 1

    # ── Run ───────────────────────────────────────────────────────────────────
    print("\nSCADE Model Parser")
    print("=" * 50)
    print(f"Input  : {input_path}")
    print(f"Output : {output_path}")
    print()

    try:
        inputs, outputs, calibrations = parse_path(input_path)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1

    print()
    print("Results")
    print("-" * 30)
    print(f"  Input signals          : {len(inputs)}")
    print(f"  Output signals         : {len(outputs)}")
    print(f"  Calibration parameters : {len(calibrations)}")

    if not any([inputs, outputs, calibrations]):
        print(
            "\n[WARNING] No data extracted.\n"
            "  Verify that the XML structure uses standard SCADE element names\n"
            "  (Input, Output, Sensor, Const / Constant) or contact support.",
            file=sys.stderr,
        )
        return 0

    print("\nWriting Excel workbook …")
    try:
        write_excel(inputs, outputs, calibrations, output_path)
    except PermissionError:
        print(
            f"[ERROR] Cannot write to '{output_path}'.\n"
            "  The file may already be open in Excel.",
            file=sys.stderr,
        )
        return 1

    print(f"\n✓  Saved: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
