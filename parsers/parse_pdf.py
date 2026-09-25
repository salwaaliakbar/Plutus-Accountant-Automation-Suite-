"""Parse a bank statement PDF.

Two-layer strategy:
 1. Barclays text-mode parser  – fast, handles Barclays multi-line descriptions.
 2. Generic table parser       – works for most other banks (HSBC, Lloyds,
                                 NatWest, Halifax, Santander, etc.) whose
                                 PDFs contain proper table cells.

If both layers return 0 transactions a ValueError is raised with a helpful
message so the API returns a readable error to the user.

Returns the same shape as parse_csv:
    {
        'date':        datetime.date,
        'description': str,
        'subcategory': str,
        'reference':   str,
        'csv_category': str,
        'money_in':    float,
        'money_out':   float,
        'balance':     float | None,
    }

Privacy: account numbers, sort codes, IBANs are NEVER returned.
"""

import re
from datetime import datetime
from pathlib import Path

import pdfplumber

from parsers._common import salvage_json_array


# ── Shared regex helpers ──────────────────────────────────────────────────────

_MONTHS = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'

# "6 Apr", "12 Apr", "6 Apr 2025"
_DATE_DMY_TEXT = re.compile(
    rf'^(\d{{1,2}}\s+{_MONTHS}(?:\s+\d{{4}})?)\s+(.*)',
    re.IGNORECASE,
)

# "01/04/2025", "01-04-2025", "01.04.2025"
_DATE_SLASH = re.compile(r'^(\d{2}[/\-\.]\d{2}[/\-\.]\d{4})\s+(.*)')

# "2025-04-01"
_DATE_ISO = re.compile(r'^(\d{4}-\d{2}-\d{2})\s+(.*)')

_AMOUNT_RE   = re.compile(r'[\d,]+\.\d{2}')
_TWO_AMOUNTS = re.compile(r'^(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$')
_ONE_AMOUNT  = re.compile(r'^(.*?)\s+([\d,]+\.\d{2})\s*$')


def _parse_amount(s: str) -> float:
    return float(str(s).replace(',', '').replace('£', '').strip())


def _parse_date_str(s: str, year_hint: int) -> 'datetime.date | None':
    """Parse a date string in any common format."""
    s = s.strip()
    for fmt in ('%d %b %Y', '%d %b'):
        try:
            d = datetime.strptime(s, fmt)
            if '%Y' not in fmt:
                d = d.replace(year=year_hint)
            return d.date()
        except ValueError:
            pass
    for sep in ('/', '-', '.'):
        for fmt in (f'%d{sep}%m{sep}%Y', f'%m{sep}%d{sep}%Y'):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    try:
        return datetime.strptime(s, '%Y-%m-%d').date()
    except ValueError:
        pass
    return None


def _try_parse_amount_cell(val) -> float | None:
    """Convert a table cell value to float, or None if not a number."""
    if val is None:
        return None
    s = str(val).strip().replace(',', '').replace('£', '').replace('(', '-').replace(')', '')
    if not s or s in ('-', '—', ''):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _infer_year(path: Path) -> int:
    """Guess statement year from the filename.

    Handles patterns like:
      'Statement 03-MAY-24 ...'  -> 2024
      'May 24.csv'               -> 2024
      'statement_2025_04.pdf'    -> 2025
      'Apr25'                    -> 2025
    """
    name = path.stem
    # 4-digit year anywhere
    m = re.search(r'20(\d{2})', name)
    if m:
        return int('20' + m.group(1))
    # DD-MMM-YY pattern (e.g. "03-MAY-24") — take the last 2-digit number after a month
    m = re.search(
        r'\d{1,2}[-/\s](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-/\s](\d{2})',
        name, re.IGNORECASE,
    )
    if m:
        return 2000 + int(m.group(1))
    # MMM-YY or MMM YY (e.g. "May 24", "Apr25")
    m = re.search(
        r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-/\s]?(\d{2})\b',
        name, re.IGNORECASE,
    )
    if m:
        return 2000 + int(m.group(1))
    # Bare 2-digit year at end of filename (e.g. "statement_24")
    m = re.search(r'[-_\s](\d{2})$', name)
    if m:
        yr = int(m.group(1))
        if 20 <= yr <= 99:
            return 2000 + yr
    return datetime.now().year


# ── Privacy filter ────────────────────────────────────────────────────────────

_SENSITIVE_RE = re.compile(
    r'\b\d{8}\b'           # 8-digit account number
    r'|\b\d{2}-\d{2}-\d{2}\b'  # sort code XX-XX-XX
    r'|IBAN\s*[:\s]+[A-Z]{2}\d{2}[A-Z0-9]+',
    re.IGNORECASE,
)


def _scrub(text: str) -> str:
    return _SENSITIVE_RE.sub('[REDACTED]', text)


# ── Layer 1: Barclays text-mode parser ───────────────────────────────────────

_BARCLAYS_SKIP = re.compile(
    r'balance brought forward|bbaallaannccee|'
    r'start balance|'
    r'date\s+description\s+money|'
    r'barclays bank|registered in england|financial services register|'
    r'authorised by the prudential|regulated by the financial|'
    r'registered office|registered no\.|'
    r'your deposit is eligible|financial services compensation|'
    r'anything wrong\?|bank of england base rate|rate effective from|'
    r'sort code|account no|swiftbic|iban|issued on|'
    r'^see optyx|^the director|stanmore|ha7 |woodcroft|'
    r'^at a glance|^\d{2}\s+[a-z]{3}\s+-\s+\d{2}\s+[a-z]{3}|'
    r'^your business|page$|continued$|^\d+$|'
    r'u commission charges|u interest paid|money in|money out|end balance',
    re.IGNORECASE,
)

_BARCLAYS_END = re.compile(
    r'balance carried forward|total payments|total receipts|helpful information',
    re.IGNORECASE,
)

_SIDEBAR_SUFFIX = re.compile(
    r'\s+u\s+.*$'
    r'|\s+Money (?:in|out)\s+.*$'
    r'|\s+End balance\s+.*$'
    r'|\s+by the Financial Services.*$'
    r'|\s+Compensation Scheme.*$',
    re.IGNORECASE,
)


def _parse_barclays(pdf, year_hint: int) -> list[dict]:
    transactions = []
    current_date = None
    prev_balance = None
    pending = None

    def flush():
        nonlocal pending
        if pending is not None:
            transactions.append(pending.copy())
            pending = None

    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        for raw_line in text.split('\n'):
            line = raw_line.strip()
            if not line:
                continue
            line = _SIDEBAR_SUFFIX.sub('', line).strip()

            if _BARCLAYS_END.search(line):
                flush()
                break

            if _BARCLAYS_SKIP.search(line):
                m = re.search(r'start balance\s+([\d,]+\.\d{2})', line, re.I)
                if m:
                    prev_balance = _parse_amount(m.group(1))
                continue

            # Date prefix?
            dm = _DATE_DMY_TEXT.match(line)
            if dm:
                date_str = dm.group(1).strip()
                line = dm.group(2).strip()
                # Inject year if not present
                if not re.search(r'\d{4}', date_str):
                    date_str = f"{date_str} {year_hint}"
                d = _parse_date_str(date_str, year_hint)
                if d:
                    current_date = d

            m2 = _TWO_AMOUNTS.match(line)
            m1 = _ONE_AMOUNT.match(line) if not m2 else None

            if m2:
                desc_part = m2.group(1).strip()
                balance = _parse_amount(m2.group(3))
                delta = balance - prev_balance if prev_balance is not None else None
                if delta is not None:
                    money_in  = round(delta, 2)  if delta >  0.001 else 0.0
                    money_out = round(-delta, 2) if delta < -0.001 else 0.0
                else:
                    money_out = _parse_amount(m2.group(2))
                    money_in = 0.0
                prev_balance = balance
                flush()
                pending = {
                    'date': current_date, 'description': _scrub(desc_part),
                    'subcategory': '', 'reference': '', 'csv_category': '',
                    'money_in': money_in, 'money_out': money_out, 'balance': balance,
                }
            elif m1:
                desc_part = m1.group(1).strip()
                if not desc_part:
                    prev_balance = _parse_amount(m1.group(2))
                    continue
                balance = _parse_amount(m1.group(2))
                delta = balance - prev_balance if prev_balance is not None else None
                if delta is not None:
                    money_in  = round(delta, 2)  if delta >  0.001 else 0.0
                    money_out = round(-delta, 2) if delta < -0.001 else 0.0
                else:
                    money_in = money_out = 0.0
                prev_balance = balance
                flush()
                pending = {
                    'date': current_date, 'description': _scrub(desc_part),
                    'subcategory': '', 'reference': '', 'csv_category': '',
                    'money_in': money_in, 'money_out': money_out, 'balance': balance,
                }
            else:
                # Fallback: find rightmost two decimal amounts
                matches = list(_AMOUNT_RE.finditer(line))
                if len(matches) >= 2:
                    balance = _parse_amount(matches[-1].group())
                    amt     = _parse_amount(matches[-2].group())
                    desc_part = line[:matches[-2].start()].strip()
                    delta = balance - prev_balance if prev_balance is not None else None
                    if delta is not None:
                        money_in  = round(delta, 2)  if delta >  0.001 else 0.0
                        money_out = round(-delta, 2) if delta < -0.001 else 0.0
                    else:
                        money_in = 0.0; money_out = amt
                    prev_balance = balance
                    flush()
                    pending = {
                        'date': current_date, 'description': _scrub(desc_part),
                        'subcategory': '', 'reference': '', 'csv_category': '',
                        'money_in': money_in, 'money_out': money_out, 'balance': balance,
                    }
                elif pending is not None and line:
                    pending['description'] = (pending['description'] + ' ' + line).strip()

    flush()
    return [
        t for t in transactions
        if t['date'] is not None and (t['money_in'] > 0 or t['money_out'] > 0)
    ]


# ── Layer 2: Generic table parser ────────────────────────────────────────────

# Column header patterns for common banks
_COL_DATE  = re.compile(r'^date$', re.I)
_COL_DESC  = re.compile(r'description|details|narrative|merchant|payee|party', re.I)
_COL_IN    = re.compile(r'paid.?in|credit|money.?in|deposits?|receipts?', re.I)
_COL_OUT   = re.compile(r'paid.?out|debit|money.?out|withdrawals?|payments?', re.I)
_COL_AMT   = re.compile(r'^amount$', re.I)
_COL_BAL   = re.compile(r'^balance', re.I)
_COL_REF   = re.compile(r'reference|ref$', re.I)
_COL_TYPE  = re.compile(r'^type$|transaction.?type', re.I)


def _match_col(headers: list[str], pattern: re.Pattern) -> int | None:
    for i, h in enumerate(headers):
        if h and pattern.search(str(h).strip()):
            return i
    return None


def _parse_generic_table(pdf, year_hint: int) -> list[dict]:
    """Extract transactions from any PDF that renders a clear table."""
    transactions = []

    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            if not table or len(table) < 2:
                continue

            # First row is usually the header
            headers = [str(c or '').strip() for c in table[0]]
            i_date = _match_col(headers, _COL_DATE)
            i_desc = _match_col(headers, _COL_DESC)
            i_in   = _match_col(headers, _COL_IN)
            i_out  = _match_col(headers, _COL_OUT)
            i_amt  = _match_col(headers, _COL_AMT)
            i_bal  = _match_col(headers, _COL_BAL)
            i_ref  = _match_col(headers, _COL_REF)
            i_type = _match_col(headers, _COL_TYPE)

            if i_date is None or i_desc is None:
                continue  # not a transaction table
            if i_in is None and i_out is None and i_amt is None:
                continue

            for row in table[1:]:
                if not row or all(c is None or str(c).strip() == '' for c in row):
                    continue

                def cell(idx):
                    if idx is None or idx >= len(row):
                        return None
                    return row[idx]

                date_val = cell(i_date)
                if not date_val or not str(date_val).strip():
                    continue
                d = _parse_date_str(str(date_val).strip(), year_hint)
                if d is None:
                    continue

                desc = _scrub(str(cell(i_desc) or '').strip())

                money_in  = _try_parse_amount_cell(cell(i_in))  or 0.0
                money_out = _try_parse_amount_cell(cell(i_out)) or 0.0
                balance   = _try_parse_amount_cell(cell(i_bal))

                # Signed amount column (positive = in, negative = out)
                if i_amt is not None and money_in == 0.0 and money_out == 0.0:
                    raw_amt = _try_parse_amount_cell(cell(i_amt))
                    if raw_amt is not None:
                        if raw_amt >= 0:
                            money_in = raw_amt
                        else:
                            money_out = -raw_amt

                if money_in == 0.0 and money_out == 0.0:
                    continue

                ref  = _scrub(str(cell(i_ref)  or '').strip())
                typ  = str(cell(i_type) or '').strip()

                transactions.append({
                    'date': d, 'description': desc,
                    'subcategory': typ, 'reference': ref, 'csv_category': '',
                    'money_in': round(money_in, 2),
                    'money_out': round(money_out, 2),
                    'balance': balance,
                })

    return transactions


# ── Layer 3: Layout (column-position) parser ─────────────────────────────────
#
# Most bank PDFs (Starling, HSBC, Lloyds, NatWest, Nationwide, Santander,
# Monzo, Tide, Metro, …) draw their statement as positioned text with no
# ruled table grid, so extract_tables() finds nothing. They do, however,
# print a header row ("Date … Paid in  Paid out  Balance") and right-align
# every amount under its column heading. This layer finds that header,
# learns each column's x-position, and assigns every amount on a row to
# In / Out / Balance by position — so Paid in and Paid out are never
# confused, whatever the bank calls them.

_LAYOUT_AMOUNT = re.compile(
    r'^(?P<pre>[^\d]{0,3}?)'
    r'(?P<num>\d{1,3}(?:,\d{3})+\.\d{2}|\d+\.\d{2})'
    r'(?P<suf>\)|CR|DR|-)?$',
    re.IGNORECASE,
)

_LAYOUT_SKIP = re.compile(
    r'\b(?:opening|closing|start|end)\s+balance\b|brought\s+forward|'
    r'carried\s+forward|^\s*(?:sub-?)?total\b',
    re.IGNORECASE,
)

_HDR_IN   = {'in', 'credit', 'credits', 'deposit', 'deposits', 'receipts'}
_HDR_OUT  = {'out', 'debit', 'debits', 'withdrawn', 'withdrawal', 'withdrawals', 'payments'}
_HDR_DESC = {'description', 'details', 'transaction', 'narrative', 'particulars',
             'payee', 'merchant', 'counterparty'}

_DATE_WORD_PATTERNS = (
    # (number of words the date spans, compiled regex over those words joined by ' ')
    (1, re.compile(r'^\d{1,2}[/\-.]\d{1,2}[/\-.](?:\d{4}|\d{2})$')),
    (1, re.compile(r'^\d{4}-\d{2}-\d{2}$')),
    (1, re.compile(rf'^\d{{1,2}}-{_MONTHS}[a-z]*-(?:\d{{4}}|\d{{2}})$', re.I)),
    (3, re.compile(rf'^\d{{1,2}}(?:st|nd|rd|th)?\s+{_MONTHS}[a-z]*\.?\s+(?:\d{{4}}|\d{{2}})$', re.I)),
    (2, re.compile(rf'^\d{{1,2}}(?:st|nd|rd|th)?\s+{_MONTHS}[a-z]*\.?$', re.I)),
)


def _norm_word(t: str) -> str:
    return re.sub(r'[^a-z]', '', t.lower())


def _layout_amount(text: str):
    """Return (value, sign_hint) for an amount token, else None.
    sign_hint: -1 for explicitly negative/debit, +1 for CR, 0 if unsigned."""
    m = _LAYOUT_AMOUNT.match(text.strip())
    if not m:
        return None
    pre, suf = m.group('pre') or '', (m.group('suf') or '').upper()
    if any(c.isalpha() for c in pre):
        return None
    val = float(m.group('num').replace(',', ''))
    sign = 0
    if '-' in pre or '−' in pre or '(' in pre or suf in (')', '-', 'DR'):
        sign = -1
    elif suf == 'CR':
        sign = 1
    return val, sign


def _layout_date(words: list, year_hint: int):
    """If the line starts with a date, return (date, n_words_used, has_year)."""
    for n, pat in _DATE_WORD_PATTERNS:
        if len(words) < n:
            continue
        s = ' '.join(w['text'] for w in words[:n])
        if not pat.match(s):
            continue
        s = re.sub(r'(\d)(st|nd|rd|th)\b', r'\1', s, flags=re.I).replace('.', ' ' if n > 1 else '.')
        s = re.sub(r'\s+', ' ', s).strip()
        for fmt in ('%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d-%m-%y', '%d.%m.%Y', '%d.%m.%y',
                    '%Y-%m-%d', '%d-%b-%Y', '%d-%b-%y', '%d %b %Y', '%d %b %y',
                    '%d %B %Y', '%d %B %y'):
            try:
                return datetime.strptime(s, fmt).date(), n, True
            except ValueError:
                pass
        for fmt in ('%d %b', '%d %B'):
            try:
                return datetime.strptime(f'{s} {year_hint}', f'{fmt} %Y').date(), n, False
            except ValueError:
                pass
    return None


def _group_lines(words: list, tol: float = 3.0) -> list:
    lines = []
    for w in sorted(words, key=lambda w: (round(w['top']), w['x0'])):
        if lines and abs(lines[-1]['top'] - w['top']) <= tol:
            lines[-1]['words'].append(w)
        else:
            lines.append({'top': w['top'], 'bottom': w['bottom'], 'words': [w]})
    for ln in lines:
        ln['words'].sort(key=lambda w: w['x0'])
        ln['bottom'] = max(w['bottom'] for w in ln['words'])
    return lines


def _detect_header(lines: list):
    """Find the statement's column-header row. Returns (cols, header_bottom)
    where cols maps 'date'/'in'/'out'/'balance'/'amount'/'type'/'desc' to
    (x0, x1), or None if this page has no header row."""
    for idx, ln in enumerate(lines):
        words = ln['words']
        toks = [_norm_word(w['text']) for w in words]
        if 'date' not in toks:
            continue
        cols = {}

        def put(key, ws):
            if key not in cols:
                cols[key] = (min(w['x0'] for w in ws), max(w['x1'] for w in ws))

        i = 0
        while i < len(words):
            t = toks[i]
            nxt = toks[i + 1] if i + 1 < len(words) else ''
            if t in ('paid', 'money') and nxt in ('in', 'out'):
                put('in' if nxt == 'in' else 'out', words[i:i + 2]); i += 2; continue
            if t in ('transaction', 'payment') and nxt == 'type':
                # "Payment type and details" (HSBC) is one combined column
                if i + 3 < len(words) and toks[i + 2] == 'and':
                    put('desc', words[i:i + 4]); i += 4; continue
                put('type', words[i:i + 2]); i += 2; continue
            if t == 'date':
                put('date', [words[i]])
            elif t in _HDR_IN:
                put('in', [words[i]])
            elif t in _HDR_OUT:
                put('out', [words[i]])
            elif t == 'balance':
                put('balance', [words[i]])
            elif t == 'amount':
                put('amount', [words[i]])
            elif t == 'type':
                put('type', [words[i]])
            elif t in _HDR_DESC:
                put('desc', [words[i]])
            i += 1

        # Multi-line headers ("END OF DAY / ACCOUNT / BALANCE"): pick up
        # money headings stacked just above or below the anchor row.
        bottom = ln['bottom']
        for other in lines[max(0, idx - 2): idx + 3]:
            if other is ln or abs(other['top'] - ln['top']) > 16:
                continue
            for w in other['words']:
                t = _norm_word(w['text'])
                if t == 'balance':
                    put('balance', [w])
                elif t in _HDR_IN and t != 'in':
                    put('in', [w])
                elif t in _HDR_OUT and t != 'out':
                    put('out', [w])
            bottom = max(bottom, other['bottom']) if other['top'] > ln['top'] else bottom

        money = [k for k in ('in', 'out', 'amount') if k in cols]
        if money and (len(money) + ('balance' in cols)) >= 2:
            return cols, bottom
    return None


def _assign_column(word: dict, cols: dict):
    """Pick the money column an amount word sits under (right-edge or centre
    alignment), or None if it is not under any money heading."""
    best, best_d = None, None
    cx = (word['x0'] + word['x1']) / 2
    for key in ('in', 'out', 'balance', 'amount'):
        if key not in cols:
            continue
        x0, x1 = cols[key]
        d = min(abs(word['x1'] - x1), abs(cx - (x0 + x1) / 2))
        if best_d is None or d < best_d:
            best, best_d = key, d
    if best is None or best_d > 40:
        return None
    return best


def _parse_layout(pdf, year_hint: int) -> list[dict]:
    rows = []          # one entry per statement line below the header
    cols = None
    for page in pdf.pages:
        words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        if not words:
            continue
        lines = _group_lines(words)
        hdr = _detect_header(lines)
        start_y = -1
        if hdr:
            cols, start_y = hdr
        if cols is None:
            continue
        money_left = min(v[0] for k, v in cols.items() if k in ('in', 'out', 'balance', 'amount'))
        text_left = max(cols['date'][1] if 'date' in cols else 0, 0)
        desc_x0 = cols['desc'][0] if 'desc' in cols else None
        prev_bottom = None
        for ln in lines:
            if ln['top'] <= start_y:
                continue
            ws = ln['words']
            line_text = ' '.join(w['text'] for w in ws)
            dt = _layout_date(ws, year_hint)
            date_limit = min([cols[k][0] - 2 for k in ('type', 'desc') if k in cols]
                             + [cols['date'][1] + 40 if 'date' in cols else 1e9])
            if dt and ws[0]['x0'] > date_limit:
                dt = None        # a date inside the description, not the date column
            used = dt[1] if dt else 0
            amounts, text_words, stray = {}, [], False
            for w in ws[used:]:
                a = _layout_amount(w['text']) if w['x0'] >= money_left - 80 else None
                key = _assign_column(w, cols) if a else None
                if key:
                    amounts[key] = a
                elif w['x1'] <= money_left + 5:
                    text_words.append(w)
                else:
                    stray = True    # text under the money columns: not a statement row
            gap = (ln['top'] - prev_bottom) if prev_bottom is not None else None
            prev_bottom = ln['bottom']
            rows.append({
                'date': dt[0] if dt else None, 'has_year': bool(dt and dt[2]),
                'amounts': amounts, 'words': text_words, 'stray': stray,
                'text': line_text, 'x0': ws[0]['x0'], 'gap': gap,
                'skip': bool(_LAYOUT_SKIP.search(line_text)),
                'text_left': text_left, 'desc_x0': desc_x0, 'cols': cols,
            })

    if not rows:
        return []

    # Does each transaction print its amounts on its FIRST line with wrapped
    # description underneath (Starling, Barclays, Santander), or on its LAST
    # line after the description (HSBC)? Decide from the dated rows.
    dated = [r for r in rows if r['date'] and not r['skip']]
    first_line_mode = sum(1 for r in dated if r['amounts']) >= sum(1 for r in dated if not r['amounts'])

    def split_text(r):
        cols_ = r['cols']
        typ, desc = [], []
        for w in r['words']:
            if 'type' in cols_ and 'desc' in cols_ and cols_['type'][0] - 5 <= w['x0'] < cols_['desc'][0] - 2:
                typ.append(w['text'])
            else:
                desc.append(w['text'])
        return ' '.join(typ), ' '.join(desc)

    txns, current_date, pending_text, pending_type = [], None, [], []
    last_month, year_bump = None, 0

    def fix_year(d, has_year):
        # Dates printed without a year ("8 Apr") on a statement that runs
        # Dec -> Jan roll into the next year. Dates with a year are kept.
        nonlocal last_month, year_bump
        if has_year:
            return d
        if last_month and d.month < last_month - 6:
            year_bump += 1
        last_month = d.month
        try:
            return d.replace(year=d.year + year_bump)
        except ValueError:
            return d

    for r in rows:
        if r['skip']:
            pending_text, pending_type = [], []
            continue
        if r['date']:
            current_date = fix_year(r['date'], r['has_year'])
        typ, desc = split_text(r)
        money = {k: v for k, v in r['amounts'].items() if k != 'balance'}

        if not money:
            if r['date'] and not first_line_mode:
                pending_text, pending_type = ([desc] if desc else []), ([typ] if typ else [])
            elif r['date'] is None and (desc or typ):
                is_cont = (r['gap'] is not None and r['gap'] < 12 and not r['stray']
                           and r['x0'] > r['text_left'] + 2)
                if first_line_mode:
                    if txns and is_cont:
                        txns[-1]['description'] = (txns[-1]['description'] + ' ' + desc).strip()
                elif is_cont or pending_text:
                    pending_text.append(desc); pending_type.append(typ)
            continue

        if current_date is None:
            continue
        money_in = money_out = 0.0
        if 'in' in money:
            money_in = money['in'][0]
        if 'out' in money:
            money_out = money['out'][0]
        amount_sign = None
        if 'amount' in money and not money_in and not money_out:
            val, sign = money['amount']
            if sign < 0:
                money_out = val
            elif sign > 0:
                money_in = val
            else:
                money_in, amount_sign = val, 0   # direction resolved from balance below
        if money_in == 0 and money_out == 0:
            continue

        bal = r['amounts'].get('balance')
        balance = (-bal[0] if bal[1] < 0 else bal[0]) if bal else None
        if first_line_mode:
            full_desc, full_type = desc, typ
        else:
            full_desc = ' '.join(t for t in pending_text + [desc] if t)
            full_type = ' '.join(t for t in pending_type + [typ] if t)
            pending_text, pending_type = [], []
        txns.append({
            'date': current_date,
            'description': _scrub(full_desc.strip()),
            'subcategory': full_type.strip(), 'reference': '', 'csv_category': '',
            'money_in': round(money_in, 2), 'money_out': round(money_out, 2),
            'balance': balance, '_unsigned': amount_sign == 0,
        })

    # Unsigned single "Amount" column: take direction from the running balance.
    prev_bal = None
    for t in txns:
        if t.pop('_unsigned') and prev_bal is not None and t['balance'] is not None:
            if t['balance'] < prev_bal - 0.001:
                t['money_out'], t['money_in'] = t['money_in'], 0.0
        if t['balance'] is not None:
            prev_bal = t['balance']
    return txns


def _reconcile(txns: list) -> tuple[int, int]:
    """Check money_in/money_out against the statement's printed running
    balance. Returns (balances_checked, mismatches). A clean statement gives
    mismatches == 0, which proves no row was dropped or put in the wrong
    Paid in / Paid out column."""
    running, checked, bad = None, 0, 0
    for t in txns:
        if running is not None:
            running = round(running + t['money_in'] - t['money_out'], 2)
        bal = t.get('balance')
        if bal is None:
            continue
        if running is None:
            running = bal
            continue
        checked += 1
        if abs(running - bal) > 0.01:
            bad += 1
            running = bal     # re-anchor so one error isn't counted repeatedly
    return checked, bad


# ── Layer 4: AI extraction (fallback for unknown layouts) ────────────────────

def _parse_with_ai(pdf, year_hint: int) -> list[dict]:
    """Ask Claude to extract transactions from statement text.

    Privacy: only lines containing money amounts are sent (this excludes the
    address block and account-details header), and account numbers, sort codes
    and IBANs are scrubbed from every line before sending.
    """
    import os
    from dotenv import load_dotenv
    load_dotenv()

    api_key = os.environ.get('ANTHROPIC_API_KEY', '')
    if not api_key or api_key.startswith('sk-ant-...'):
        return []

    # Collect candidate transaction lines: only lines with an amount on them
    lines = []
    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        for raw in text.split('\n'):
            line = raw.strip()
            if line and _AMOUNT_RE.search(line):
                lines.append(_scrub(line))
    if not lines:
        return []

    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    model = os.environ.get('CATEGORISE_MODEL', 'claude-sonnet-5')

    system = (
        "You extract transactions from UK bank statement text. You are given a "
        "NUMBERED list of statement lines. Return ONLY a JSON array with EXACTLY "
        "one entry per input line, in the same order — this is critical: do not "
        "skip, merge or add lines, even if a line looks like a duplicate or its "
        "running balance happens to equal a total shown elsewhere (e.g. in a "
        "summary box) — that is a coincidence of the running total, not a reason "
        "to drop it.\n"
        "For a line that IS a real transaction, its entry is:\n"
        '{"date":"YYYY-MM-DD","description":"...","money_in":0.0,'
        '"money_out":0.0,"balance":null_or_number}\n'
        "For a line that is NOT a transaction (a header, column title, total, "
        "brought/carried-forward line, footer, or unrelated text), its entry is "
        "the JSON value null — do not omit it from the array, output null in its "
        "place so the array length always matches the number of input lines. "
        "A row with a date and a transaction type (e.g. FASTER PAYMENT, DIRECT "
        "CREDIT, CARD PAYMENT, TRANSFER, DIRECT DEBIT, ONLINE PAYMENT) is always "
        "a real transaction, never null. "
        "Dates without a year belong to the statement year given by the user. "
        "money_in for credits, money_out for debits (both positive numbers). "
        "Keep the description as written, minus any '[REDACTED]' markers."
    )

    # Transactions are keyed by the statement LINE they came from, not by
    # date/description/amount: two genuine identical payments on one day
    # (e.g. two £3.40 fares) sit on two different lines and are both kept,
    # while the same line seen twice via chunk overlap collapses to one.
    by_line = {}        # line index -> transaction
    unplaced = []       # (approx line index, transaction) that couldn't be tied to a line

    def amount_in_line(amt: float, line: str) -> bool:
        # digit boundaries so 5.00 doesn't match a line showing 15.00
        return any(re.search(rf'(?<![\d,.]){re.escape(s)}(?!\d)', line)
                   for s in {f"{amt:,.2f}", f"{amt:.2f}"})

    def to_txn(row):
        try:
            d = datetime.strptime(str(row['date']), '%Y-%m-%d').date()
            money_in  = float(row.get('money_in') or 0)
            money_out = float(row.get('money_out') or 0)
            bal = row.get('balance')
            balance = float(bal) if bal is not None else None
            desc = str(row.get('description') or '').replace('[REDACTED]', '').strip()
        except (KeyError, TypeError, ValueError, AttributeError):
            return None
        if money_in <= 0 and money_out <= 0:
            return None
        return {
            'date': d, 'description': desc,
            'subcategory': '', 'reference': '', 'csv_category': '',
            'money_in': round(money_in, 2),
            'money_out': round(money_out, 2),
            'balance': balance,
        }

    def ask_chunk(start: int, chunk: list) -> None:
        numbered = '\n'.join(f"{i + 1}. {l}" for i, l in enumerate(chunk))
        msg = (
            f"Statement year: {year_hint}\n"
            f"Extract from these {len(chunk)} numbered lines — return exactly "
            f"{len(chunk)} array entries, one per line:\n" + numbered
        )
        response = client.messages.create(
            model=model, max_tokens=16000, system=system,
            messages=[{'role': 'user', 'content': msg}],
        )
        text = ''.join(
            b.text for b in response.content if getattr(b, 'type', '') == 'text'
        ).strip()
        rows = salvage_json_array(text)
        if not rows:
            return
        aligned = len(rows) == len(chunk)
        claimed = set()     # lines already matched in THIS pass
        for j, row in enumerate(rows):
            t = to_txn(row) if isinstance(row, dict) else None
            if t is None:
                continue
            amt = t['money_in'] or t['money_out']
            # 1:1 reply (as asked): entry j is line j, if its amount is on it.
            idx = start + j if aligned and amount_in_line(amt, chunk[j]) else None
            if idx is None:
                # Reply skipped/merged lines: tie it to the first line in this
                # chunk that carries this amount and isn't matched yet.
                idx = next((start + k for k, l in enumerate(chunk)
                            if start + k not in claimed and amount_in_line(amt, l)), None)
            if idx is None:
                unplaced.append((start + j, t))
                continue
            claimed.add(idx)
            by_line.setdefault(idx, t)      # first pass to read a line wins

    # Kept small: Claude's extended-thinking budget is drawn from the same
    # max_tokens limit as the reply, and how much it uses varies per call —
    # a large chunk risks the JSON array getting cut off before it closes
    # (see salvage_json_array). A smaller chunk keeps the expected output
    # comfortably inside the budget even when thinking eats into it.
    #
    # Chunks overlap by OVERLAP lines: LLM extraction occasionally drops a
    # line by chance (independent of truncation), most visibly right at a
    # chunk boundary. Re-sending the tail of one chunk as the head of the
    # next gives a second chance at each line; keying results by line
    # index (by_line) collapses the resulting repeats back down.
    CHUNK = 40
    OVERLAP = 4
    STEP = CHUNK - OVERLAP
    for start in range(0, len(lines), STEP):
        ask_chunk(start, lines[start:start + CHUNK])
        if start + CHUNK >= len(lines):
            break

    # The very last line has no following chunk to give it a second chance
    # via overlap, so it is the one spot a one-off drop would go uncaught —
    # re-check the tail once more.
    if len(lines) > OVERLAP:
        tail = min(CHUNK, len(lines))
        ask_chunk(len(lines) - tail, lines[-tail:])

    if unplaced:
        print(f"  [parse_pdf] WARNING: {len(unplaced)} AI-extracted transaction(s) "
              "could not be matched to a statement line — kept, please review.")
    placed = sorted(list(by_line.items()) + unplaced, key=lambda p: p[0])
    return [t for _, t in placed]


# ── Public entry point ────────────────────────────────────────────────────────

def _order(txns: list) -> list:
    """Ascending by date, keeping the statement's own order within a day.
    Newest-first statements are reversed first so same-day rows don't end
    up backwards (which would break the running-balance check)."""
    if len(txns) > 1 and txns[0]['date'] > txns[-1]['date']:
        txns = list(reversed(txns))
    return sorted(txns, key=lambda t: t['date'])


def parse_pdf(path, year_hint: int = None) -> list[dict]:
    """Parse a bank statement PDF and return transactions.

    Three layers: Barclays text parser -> generic table parser -> AI extraction.
    Raises ValueError if none succeed.

    year_hint: the calendar year (e.g. 2025). Auto-detected from the filename
               if not supplied.
    """
    path = Path(path)
    if year_hint is None:
        year_hint = _infer_year(path)

    with pdfplumber.open(path) as pdf:
        # Layer 1: Barclays text-mode (only if it actually is a Barclays statement)
        first_text = ' '.join(
            (p.extract_text() or '') for p in pdf.pages[:2]
        ).lower()
        if 'barclays' in first_text:
            txns = _parse_barclays(pdf, year_hint)
            if txns:
                print(f"  [parse_pdf] Barclays parser: {len(txns)} transactions from {path.name}")
                return sorted(txns, key=lambda t: t['date'])

        # Layers 2 & 3: ruled-table and column-position parsers. Both run;
        # the one whose Paid in / Paid out reconcile with the printed
        # running balance wins (then the one that found more rows).
        candidates = []
        for name, fn in (('Generic table', _parse_generic_table), ('Layout', _parse_layout)):
            try:
                txns = fn(pdf, year_hint)
            except Exception as exc:     # a layout quirk must not kill the upload
                print(f"  [parse_pdf] {name} parser failed: {exc}")
                continue
            if txns:
                checked, bad = _reconcile(txns)
                candidates.append((bad, -len(txns), name, txns, checked))
        if candidates:
            bad, _, name, txns, checked = min(candidates, key=lambda c: (c[0], c[1]))
            print(f"  [parse_pdf] {name} parser: {len(txns)} transactions from {path.name} "
                  f"(balance check: {checked - bad}/{checked} rows reconcile)")
            return _order(txns)

        # Layer 4: AI extraction (unknown layout)
        print(f"  [parse_pdf] Standard parsers found nothing — trying AI extraction...")
        txns = _parse_with_ai(pdf, year_hint)
        if txns:
            print(f"  [parse_pdf] AI parser: {len(txns)} transactions from {path.name}")
            return _order(txns)

    raise ValueError(
        f"Could not read any transactions from '{path.name}'. "
        "The PDF may be a scanned image (no selectable text). "
        "Please export a CSV from your online banking instead — "
        "most portals offer a 'Download transactions (CSV)' option. "
        f"(Year hint: {year_hint})"
    )
