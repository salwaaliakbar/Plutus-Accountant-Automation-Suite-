"""Shared column-detection & row-parsing logic for CSV and XLSX bank
statements, so that the same underlying data produces the same output
regardless of which format the client happened to export.

Handles:
  - Signed Amount column  (Barclays standard)
  - Separate Paid in / Paid out columns  (most other banks)
  - Native XLSX date/number cells as well as formatted strings
  - Comma-formatted numbers  (1,000.00)
  - Newest-first ordering — always returns ascending
"""

import re
from datetime import date, datetime


def clean(raw) -> str:
    if raw is None:
        return ''
    return re.sub(r'\s{2,}', ' ', str(raw).replace('\t', ' ')).strip()


def cell_str(val) -> str:
    """Stringify a cell for text fields (reference, type, …), coping with
    XLSX cells that come back as int/float instead of str."""
    if val is None:
        return ''
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def parse_amount(raw):
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    raw = str(raw)
    if not raw.strip():
        return None
    cleaned = re.sub(r'[£,\s]', '', raw.strip()).replace('(', '-').replace(')', '')
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_datetime(raw):
    """Return a datetime, accepting either a native date/datetime (XLSX
    cells) or a formatted string (CSV cells)."""
    if isinstance(raw, datetime):
        return raw
    if isinstance(raw, date):
        return datetime(raw.year, raw.month, raw.day)
    raw = str(raw).strip()
    if not raw:
        return None
    for fmt in ('%d/%m/%Y %H:%M', '%d/%m/%Y', '%m/%d/%Y',
                '%Y-%m-%d', '%d-%m-%Y', '%d %b %Y', '%d/%m/%y'):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def find_col(headers: list, *candidates: str):
    """Match a header by name, case-insensitively. Returns None (never a
    guessed fallback column) when nothing matches — better to leave a
    field blank than to silently fill it with the wrong column's data."""
    norm = {str(h).strip().lower(): h for h in headers if h is not None}
    for c in candidates:
        found = norm.get(c.lower())
        if found is not None:
            return found
    return None


def parse_rows(rows: list, headers: list) -> list:
    """rows: list of {header: cell_value} dicts. cell_value may be a plain
    string (CSV) or a native str/int/float/date/datetime (XLSX).

    Returns the common transaction shape used across all parsers:
        {date, timestamp, description, subcategory, reference,
         csv_category, money_in, money_out, balance,
         from, to, status, tag}
    """
    date_col = find_col(headers, 'Date', 'Transaction Date', 'Value Date')
    if not date_col:
        raise ValueError(f"No Date column found. Headers: {headers}")

    desc_col = find_col(headers,
        'Memo', 'Counter Party', 'Description', 'Transaction description',
        'Details', 'Narrative', 'Payee', 'From',
    )
    ref_col      = find_col(headers, 'Reference', 'Ref', 'Transaction ID')
    type_col     = find_col(headers, 'Type', 'Transaction Type', 'Transaction type')
    csv_cat_col  = find_col(headers, 'Spending Category', 'Category name', 'Category')
    bal_col      = find_col(headers, 'Balance', 'Balance (GBP)', 'Balance (£)', 'Running Balance')

    in_col  = find_col(headers, 'Paid in', 'Paid In', 'IN', 'In', 'In (£)', 'In (GBP)', 'Credit', 'Money in', 'Money In')
    out_col = find_col(headers, 'Paid out', 'Paid Out', 'OUT', 'Out', 'Out (£)', 'Out (GBP)', 'Debit', 'Money out', 'Money Out')
    amt_col = find_col(headers, 'Amount', 'Amount (GBP)', 'Amount (£)')

    from_col   = find_col(headers, 'From')
    to_col     = find_col(headers, 'To')
    status_col = find_col(headers, 'Status')
    tag_col    = find_col(headers, 'Tag 1', 'Tag')

    transactions = []
    for row in rows:
        date_raw = row.get(date_col)
        if date_raw is None or (isinstance(date_raw, str) and not date_raw.strip()):
            continue
        dt = parse_datetime(date_raw)
        if dt is None:
            continue
        d = dt.date()
        timestamp = dt if (dt.hour or dt.minute or dt.second) else None

        description  = clean(row.get(desc_col))       if desc_col    else ''
        reference    = cell_str(row.get(ref_col))      if ref_col     else ''
        subcategory  = cell_str(row.get(type_col))     if type_col    else ''
        csv_category = cell_str(row.get(csv_cat_col))  if csv_cat_col else ''

        # Strip leading apostrophe Excel inserts for text-prefix formatting
        reference = reference.lstrip("'")

        money_in = money_out = 0.0
        if in_col or out_col:
            money_in  = parse_amount(row.get(in_col))  or 0.0
            money_out = parse_amount(row.get(out_col)) or 0.0
        elif amt_col:
            amt = parse_amount(row.get(amt_col))
            if amt is None:
                continue
            if amt >= 0:
                money_in = amt
            else:
                money_out = abs(amt)
        else:
            continue  # no usable money column

        if money_in == 0 and money_out == 0:
            continue  # blank / header row

        balance = parse_amount(row.get(bal_col)) if bal_col else None

        transactions.append({
            'date':         d,
            'timestamp':    timestamp,
            'description':  description,
            'subcategory':  subcategory,
            'reference':    reference,
            'csv_category': csv_category,
            'money_in':     money_in,
            'money_out':    money_out,
            'balance':      balance,
            'from':         cell_str(row.get(from_col))   if from_col   else '',
            'to':           cell_str(row.get(to_col))     if to_col     else '',
            'status':       cell_str(row.get(status_col)) if status_col else '',
            'tag':          cell_str(row.get(tag_col))    if tag_col    else '',
        })

    # Sort ascending by date (some bank exports are newest-first)
    transactions.sort(key=lambda t: t['date'])
    return transactions
