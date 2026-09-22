"""Universal bank statement CSV parser.

Auto-detects column layout from headers. Handles:
  - Signed Amount column  (Barclays standard, Matthew Farris)
  - Separate Paid in / Paid out columns  (Chido Hove, Ahmed Ibrahim, Shaifa Remtulla)
  - Date-time stamps  (DD/MM/YYYY HH:MM)
  - Comma-formatted numbers  (1,000.00)
  - Newest-first ordering — always returns ascending
  - Multiple encodings  (utf-8-sig, utf-8, latin-1, cp1252)

Column detection and row parsing are shared with parse_xls.py (see
parsers/_common.py) so that a CSV and an XLSX export of the same statement
produce identical output.
"""

import csv
from pathlib import Path

from parsers._common import parse_rows


def _open_csv(path: Path):
    for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
        try:
            fh = open(path, newline='', encoding=enc)
            fh.read(2048)
            fh.seek(0)
            return fh, enc
        except UnicodeDecodeError:
            try:
                fh.close()
            except Exception:
                pass
    raise ValueError(f"Cannot decode {path.name} — tried utf-8-sig, utf-8, latin-1, cp1252")


def parse_csv(path) -> list[dict]:
    path = Path(path)
    fh, _ = _open_csv(path)
    try:
        reader = csv.DictReader(fh)
        rows = list(reader)
        raw_headers = list(reader.fieldnames or [])
    finally:
        fh.close()

    if not rows:
        return []

    # Strip whitespace from all headers and row keys
    headers = [h.strip() for h in raw_headers]
    rows = [{k.strip(): v for k, v in row.items()} for row in rows]

    return parse_rows(rows, headers)
