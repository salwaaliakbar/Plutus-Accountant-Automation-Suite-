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

import json
import re
from datetime import date, datetime, timedelta


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
    """Parse an amount cell. Handles '£1,000.00', '-5.00', '(5.00)',
    '5.00-', '5.00 DR' / '5.00 CR', 'GBP 5.00' and the unicode minus sign."""
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip()
    if not s:
        return None
    neg = False
    m = re.search(r'\s*\b(CR|DR)\.?$', s, re.IGNORECASE)
    if m:
        neg = m.group(1).upper() == 'DR'
        s = s[:m.start()]
    s = s.replace('−', '-')
    s = re.sub(r'GBP|[£$€,\s]', '', s, flags=re.IGNORECASE)
    if s.startswith('(') and s.endswith(')'):
        neg, s = True, s[1:-1]
    if s.endswith('-'):
        neg, s = True, s[:-1]
    try:
        val = float(s)
    except ValueError:
        return None
    return -abs(val) if neg else val


_DATE_FORMATS = (
    '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y', '%d/%m/%y',
    '%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
    '%d-%m-%Y', '%d-%m-%y', '%d.%m.%Y', '%d.%m.%y',
    '%d %b %Y', '%d %B %Y', '%d %b %y', '%d-%b-%Y', '%d-%b-%y', '%d/%b/%Y',
    '%b %d, %Y', '%B %d, %Y',
    '%m/%d/%Y',   # last: US order only when day-first is impossible
)


def parse_datetime(raw):
    """Return a datetime, accepting a native date/datetime (XLSX cells), an
    Excel serial number, or a formatted string (CSV cells)."""
    if isinstance(raw, datetime):
        return raw
    if isinstance(raw, date):
        return datetime(raw.year, raw.month, raw.day)
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        # Excel serial date (a date cell formatted as General / Number)
        if 20000 < raw < 80000:
            return datetime(1899, 12, 30) + timedelta(days=float(raw))
        return None
    s = str(raw).strip()
    if not s:
        return None
    s = re.sub(r'(\d)(st|nd|rd|th)\b', r'\1', s, flags=re.IGNORECASE)
    s = re.sub(r'(:\d\d)\.\d+$', r'\1', s)   # fractional seconds
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def norm_header(h) -> str:
    """'Paid In (£)' / ' Paid in ' / 'Amount (GBP)' -> comparable key."""
    s = str(h).lower().replace('£', ' ')
    s = re.sub(r'\((?:gbp|\s*)\)|\bgbp\b', ' ', s)
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    return s.strip()


def find_col(headers: list, *candidates: str):
    """Match a header by name, ignoring case, spacing, punctuation and
    currency markers. Returns None (never a guessed fallback column) when
    nothing matches."""
    norm = {}
    for h in headers:
        if h is not None:
            norm.setdefault(norm_header(h), h)
    for c in candidates:
        found = norm.get(norm_header(c))
        if found is not None:
            return found
    return None


# Header names seen across UK bank exports (Barclays, Starling, Monzo, HSBC,
# Lloyds, NatWest, Santander, Nationwide, Tide, Revolut, Metro, Wise, …).
DATE_HEADERS = ('Date', 'Transaction Date', 'Value Date', 'Posting Date', 'Posted Date',
                'Date Posted', 'Booking Date', 'Completed Date', 'Started Date', 'Created')
DESC_HEADERS = ('Memo', 'Counter Party', 'Counterparty', 'Description',
                'Transaction description', 'Transaction Details', 'Details',
                'Narrative', 'Particulars', 'Payee', 'Payee Name', 'Name',
                'Merchant', 'Merchant Name', 'Beneficiary', 'Transaction')
REF_HEADERS  = ('Reference', 'Ref', 'Payment Reference', 'Transaction ID', 'Notes')
TYPE_HEADERS = ('Type', 'Transaction Type', 'Payment Type')
CAT_HEADERS  = ('Spending Category', 'Category name', 'Category')
BAL_HEADERS  = ('Balance', 'Running Balance', 'Account Balance')
IN_HEADERS   = ('Paid in', 'Money in', 'In', 'Credit', 'Credits', 'Credit Amount',
                'Receipts', 'Deposits', 'Paid In Amount')
OUT_HEADERS  = ('Paid out', 'Money out', 'Out', 'Debit', 'Debits', 'Debit Amount',
                'Withdrawn', 'Withdrawals', 'Payments', 'Paid Out Amount')
AMT_HEADERS  = ('Amount', 'Value', 'Transaction Amount', 'Net Amount')


def find_header_row(rows: list, max_scan: int = 30):
    """Many bank exports put title / account-summary lines above the real
    column headers. Return the index of the first row that has a date
    header plus a money header (in/out/amount), or None."""
    for i, row in enumerate(rows[:max_scan]):
        hdrs = [c for c in row if c not in (None, '')]
        if not hdrs:
            continue
        if find_col(hdrs, *DATE_HEADERS) and (
            find_col(hdrs, *IN_HEADERS) or find_col(hdrs, *OUT_HEADERS)
            or find_col(hdrs, *AMT_HEADERS)
        ):
            return i
    return None


def rows_to_dicts(rows: list):
    """Locate the header row inside raw rows (lists) and return
    (headers, [{header: value}]). Falls back to row 0 as the header."""
    if not rows:
        return [], []
    hi = find_header_row(rows)
    if hi is None:
        hi = 0
    headers, seen = [], {}
    for h in rows[hi]:
        h = '' if h is None else str(h).strip()
        if h and h in seen:                 # keep duplicate headers distinct
            seen[h] += 1
            h = f'{h}.{seen[h]}'
        else:
            seen[h] = 0
        headers.append(h)
    out = []
    for row in rows[hi + 1:]:
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        out.append({h: v for h, v in zip(headers, row) if h})
    return [h for h in headers if h], out


def _guess_desc_col(rows: list, headers: list, taken: set):
    """No recognised description header: use the text column with the most
    distinct non-numeric, non-date values — a best guess beats handing the
    accountant rows with an amount and nothing else."""
    best, best_n = None, 0
    for h in headers:
        if h in taken:
            continue
        texts = {str(v).strip() for v in (row.get(h) for row in rows[:200])
                 if isinstance(v, str) and v.strip()
                 and parse_amount(v) is None and parse_datetime(v) is None}
        if len(texts) > best_n:
            best, best_n = h, len(texts)
    return best if best_n >= 2 else None


def parse_rows(rows: list, headers: list) -> list:
    """rows: list of {header: cell_value} dicts. cell_value may be a plain
    string (CSV) or a native str/int/float/date/datetime (XLSX).

    Returns the common transaction shape used across all parsers:
        {date, timestamp, description, subcategory, reference,
         csv_category, money_in, money_out, balance,
         from, to, status, tag}
    """
    date_col = find_col(headers, *DATE_HEADERS)
    if not date_col:
        raise ValueError(
            "Could not find a Date column in this file. Columns found: "
            f"{', '.join(str(h) for h in headers) or '(none)'}"
        )

    desc_col    = find_col(headers, *DESC_HEADERS)
    ref_col     = find_col(headers, *REF_HEADERS)
    type_col    = find_col(headers, *TYPE_HEADERS)
    csv_cat_col = find_col(headers, *CAT_HEADERS)
    bal_col     = find_col(headers, *BAL_HEADERS)
    in_col      = find_col(headers, *IN_HEADERS)
    out_col     = find_col(headers, *OUT_HEADERS)
    amt_col     = find_col(headers, *AMT_HEADERS)
    time_col    = find_col(headers, 'Time')
    from_col    = find_col(headers, 'From')
    to_col      = find_col(headers, 'To')
    status_col  = find_col(headers, 'Status')
    tag_col     = find_col(headers, 'Tag 1', 'Tag')

    if not (in_col or out_col or amt_col):
        raise ValueError(
            "Could not find Paid in / Paid out or Amount columns in this file. "
            f"Columns found: {', '.join(str(h) for h in headers)}"
        )
    if desc_col is None:
        taken = {c for c in (date_col, ref_col, type_col, csv_cat_col, bal_col,
                             in_col, out_col, amt_col, time_col, status_col) if c}
        desc_col = _guess_desc_col(rows, headers, taken)
        if desc_col:
            print(f"  [parse] No description header recognised — using column '{desc_col}'")

    transactions = []
    skipped = []
    for n, row in enumerate(rows):
        date_raw = row.get(date_col)
        if date_raw is None or (isinstance(date_raw, str) and not date_raw.strip()):
            continue
        dt = parse_datetime(date_raw)
        if dt is None:
            skipped.append(f"row {n + 1}: unreadable date {date_raw!r}")
            continue
        if time_col and not (dt.hour or dt.minute):
            t = parse_datetime(f"{dt:%d/%m/%Y} {cell_str(row.get(time_col))[:5]}")
            if t:
                dt = t
        d = dt.date()
        timestamp = dt if (dt.hour or dt.minute or dt.second) else None

        description  = clean(row.get(desc_col))       if desc_col    else ''
        reference    = cell_str(row.get(ref_col))      if ref_col     else ''
        subcategory  = cell_str(row.get(type_col))     if type_col    else ''
        csv_category = cell_str(row.get(csv_cat_col))  if csv_cat_col else ''

        # Strip leading apostrophe Excel inserts for text-prefix formatting
        reference = reference.lstrip("'")

        money_in = money_out = 0.0
        raw_in  = parse_amount(row.get(in_col))  if in_col  else None
        raw_out = parse_amount(row.get(out_col)) if out_col else None
        if raw_in is not None or raw_out is not None:
            # some banks print the out column as negative — magnitudes count
            money_in, money_out = abs(raw_in or 0.0), abs(raw_out or 0.0)
        elif amt_col:
            amt = parse_amount(row.get(amt_col))
            if amt is None:
                if row.get(amt_col) not in (None, ''):
                    skipped.append(f"row {n + 1}: unreadable amount {row.get(amt_col)!r}")
                continue
            if amt >= 0:
                money_in = amt
            else:
                money_out = abs(amt)

        if money_in == 0 and money_out == 0:
            continue  # zero-value / blank row

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

    if skipped:
        print(f"  [parse] WARNING: skipped {len(skipped)} row(s) that could not be "
              f"read: {skipped[:5]}")

    # Ascending by date, keeping the file's own order within a day. Exports
    # that are newest-first are reversed first, so same-day rows aren't left
    # backwards by the stable sort — the same rows then come out in the same
    # order whichever format the statement was exported in.
    if len(transactions) > 1 and transactions[0]['date'] > transactions[-1]['date']:
        transactions.reverse()
    transactions.sort(key=lambda t: t['date'])
    return transactions


def salvage_json_array(text: str):
    """Parse a JSON array out of an LLM response, tolerating truncation.

    Claude's extended-thinking budget is drawn from the same max_tokens
    limit as the reply, so how much thinking a request triggers (and
    therefore how much room is left for output) varies per call — a
    generous max_tokens is not a guarantee the array will finish. If the
    response got cut off mid-element, a naive `json.loads` on the whole
    thing fails and the entire batch would otherwise be silently
    discarded. Instead, walk the array and keep every complete leading
    element, dropping only the truncated tail one.

    Returns the parsed list, or None if nothing could be recovered.
    """
    start = text.find('[')
    if start == -1:
        return None
    body = text[start:]

    try:
        return json.loads(body)
    except json.JSONDecodeError:
        pass

    depth = 0
    in_str = False
    esc = False
    top_level_commas = []
    for i, ch in enumerate(body):
        if in_str:
            if esc:
                esc = False
            elif ch == '\\':
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch in '{[':
            depth += 1
        elif ch in '}]':
            depth -= 1
        elif ch == ',' and depth == 1:
            top_level_commas.append(i)

    if not top_level_commas:
        return None

    candidate = body[:top_level_commas[-1]] + ']'
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        return None
