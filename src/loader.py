"""Reads the employee list (CSV or Excel) and returns the unique employee numbers in file order.

The file has no fixed layout: a typical payroll export is two columns with no header
(`PAYGROUP,1234`). The employee-number column defaults to the right-most column whose values are
mostly whole numbers; the user can pick another one.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Sheet:
    path: Path
    rows: list[list[str]]
    columns: list[str] = field(default_factory=list)


def _clean(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)  # Excel hands numbers back as floats: 1234.0 -> 1234
    return str(v).strip()


def read(path: str | Path) -> Sheet:
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        rows = [[_clean(c) for c in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
        wb.close()
    else:
        raw = path.read_bytes()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1252", errors="replace")
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        rows = [[_clean(c) for c in r] for r in csv.reader(text.splitlines(), dialect)]
    rows = [r for r in rows if any(r)]
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]

    header = None
    if len(rows) > 1 and not any(_is_number(v) for v in rows[0]):
        header, rows = rows[0], rows[1:]
    columns = [f"Column {_col_letter(i)}" + (f" ({header[i]})" if header and header[i] else "")
               for i in range(width)]
    return Sheet(path, rows, columns)


def _col_letter(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _is_number(v: str) -> bool:
    return normalize(v).isdigit()


def guess_column(sheet: Sheet) -> int:
    """Right-most column where at least half of the non-blank values are whole numbers."""
    for i in reversed(range(len(sheet.columns))):
        vals = [r[i] for r in sheet.rows if r[i]]
        if vals and sum(_is_number(v) for v in vals) >= len(vals) / 2:
            return i
    return max(len(sheet.columns) - 1, 0)


def normalize(v: str) -> str:
    """'  01234 ', '1234.0' -> '1234' (Spectrum's employee codes are numbers, stored padded)."""
    v = v.strip()
    if v.endswith(".0"):
        v = v[:-2]
    return str(int(v)) if v.isdigit() else v


def unique_numbers(sheet: Sheet, col: int) -> tuple[list[str], int]:
    """(unique employee numbers in first-seen order, number of non-blank rows read)."""
    seen = {}  # dict as an insertion-ordered set
    total = 0
    for r in sheet.rows:
        v = normalize(r[col]) if col < len(r) else ""
        if v:
            total += 1
            seen.setdefault(v, None)
    return list(seen), total
