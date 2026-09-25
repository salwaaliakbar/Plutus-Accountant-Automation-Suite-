"""Universal bank statement XLSX parser.

Reads the first worksheet directly with openpyxl (keeping native date and
number cell types, rather than round-tripping everything through strings)
and shares column-detection / row-parsing with parse_csv.py (see
parsers/_common.py), so an XLSX and a CSV export of the same statement
produce identical output — handling separate Paid in / Paid out columns
just as well as a single signed Amount column (Barclays standard).
"""

from pathlib import Path

import openpyxl

from parsers._common import find_header_row, parse_rows, rows_to_dicts


def parse_xls(path) -> list[dict]:
    path = Path(path)

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        # Use the first sheet that actually has a transaction header row —
        # exports sometimes lead with a summary / cover sheet.
        sheets = [list(ws.iter_rows(values_only=True)) for ws in wb.worksheets]
    finally:
        wb.close()

    chosen = next((rows for rows in sheets if find_header_row(rows) is not None),
                  sheets[0] if sheets else [])
    headers, dict_rows = rows_to_dicts(chosen)
    if not dict_rows:
        return []
    return parse_rows(dict_rows, headers)
