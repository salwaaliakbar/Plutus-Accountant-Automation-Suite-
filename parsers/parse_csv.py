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
import io
from pathlib import Path

from parsers._common import parse_rows, rows_to_dicts


def _read_text(path: Path) -> str:
    # Decode the WHOLE file per attempt: a stray £ past the first few KB
    # used to pass the probe and then crash half-way through reading.
    raw = path.read_bytes()
    for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode {path.name} — tried utf-8, cp1252, latin-1")


def parse_csv(path) -> list[dict]:
    path = Path(path)
    text = _read_text(path)
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=',;	|')
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(io.StringIO(text), dialect))
    if not rows:
        return []

    # The header row isn't always row 1 — some banks put account details
    # above it. rows_to_dicts finds it (shared with parse_xls).
    headers, dict_rows = rows_to_dicts(rows)
    if not dict_rows:
        return []
    return parse_rows(dict_rows, headers)
