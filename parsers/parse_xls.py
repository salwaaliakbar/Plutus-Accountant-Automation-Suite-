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

from parsers._common import parse_rows


def parse_xls(path) -> list[dict]:
    path = Path(path)

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb.worksheets[0]
        rows_iter = ws.iter_rows(values_only=True)
        try:
            raw_headers = next(rows_iter)
        except StopIteration:
            return []

        headers = [str(h).strip() if h is not None else None for h in raw_headers]

        rows = []
        for values in rows_iter:
            if all(v is None for v in values):
                continue
            row = {}
            for h, v in zip(headers, values):
                if h is None:
                    continue
                row[h] = v
            rows.append(row)
    finally:
        wb.close()

    if not rows:
        return []

    return parse_rows(rows, [h for h in headers if h is not None])
