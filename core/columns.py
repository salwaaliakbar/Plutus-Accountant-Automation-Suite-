"""Work out which template column plays which role.

Every accountant names the columns of their working papers differently
('UC category', 'Category  UC', 'UC desc'; 'SA', 'Year', 'Personal'; 'Spending',
'Category Bank', 'Spending Category'; 'Reference', 'Description 2' ...). A fixed
list of names therefore fails on the next client's template, and a column the
program does not recognise is silently left blank.

resolve_columns() finds each role in three steps, stopping at the first that
works:
  1. the exact names the program has always known (so existing templates
     behave exactly as before);
  2. the words inside the header ('uc' + 'category' / 'desc', 'bank' or
     'spending' + 'category', a trailing '2' on a description ...);
  3. the contents of the column (year labels such as FY24 or 23/24).

It is shared by everything that reads a template — the writer, the
client-history examples for the AI, the duplicate check and the feedback
learner — so they can never disagree about where the category column is.
"""

import re

from parsers._common import DATE_HEADERS, norm_header

# ── Exact names (the original behaviour) ──────────────────────────────────────
DESC_COLS = ('Counter Party', 'Counterparty', 'Description', 'Transaction description',
             'Transaction Details', 'Details', 'Memo', 'Narrative', 'Particulars',
             'Payee', 'Name', 'Merchant')
IN_COLS = ('Paid in', 'Money in', 'In', 'Credit', 'Credits', 'Deposits')
OUT_COLS = ('Paid out', 'Money out', 'Out', 'Debit', 'Debits', 'Payments',
            'Withdrawals', 'Withdrawn')
UC_COLS = ('UC category', 'UC Category', 'UC')
CORRECTION_COLS = ('Correct category', 'Correction', 'Accountant category')

_SA_NAMES = ('SA',)
_ACC_NAMES = ('ACC', 'AC')
_FY_NAMES = ('FY',)
_BANK_CAT_NAMES = ('Spending Category', 'Category name', 'Spending', 'Category Bank',
                   'Bank Category', 'Bank Cat')
_REF_NAMES = ('Reference', 'Ref', 'Description 2', 'Reference 2', 'Details 2',
              'Payment Reference')

_YEAR_LABEL = re.compile(
    r'^(?:(?:FY|SA|YE|AC|ACC)\s?\d{2,4}(?:\s?[/\-.]\s?\d{2,4})?'   # FY24, SA24/25, YE25
    r'|\d{2}\s?[/\-]\s?\d{2}'                                       # 24/25
    r'|(?:19|20)\d{2})$',                                           # 2025
    re.IGNORECASE)


def read_headers(ws) -> dict:
    """{normalised header: first column index} from row 1."""
    row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    hdrs = {}
    for i, v in enumerate(row1):
        # a punctuation-only header ('?', '*', '£') has no name to match on
        if v is not None and norm_header(v):
            hdrs.setdefault(norm_header(v), i + 1)
    return hdrs


def named(hdrs: dict, *names: str):
    for n in names:
        key = norm_header(n)
        idx = hdrs.get(key) if key else None
        if idx is not None:
            return idx
    return None


def last_data_row(ws, key_cols) -> int:
    """Last row where any of the key columns holds a value — formatted but
    empty rows below the data (openpyxl counts them in max_row) don't count."""
    last = 1
    for r in range(2, ws.max_row + 1):
        for c in key_cols:
            if c and ws.cell(r, c).value not in (None, ''):
                last = r
                break
    return last


def _year_style(ws, col: int, last_row: int) -> str:
    """Which year column a content-detected one is: 'fy' (FY25 / 2025),
    'sa' (SA24/25) or 'acc' (24/25)."""
    votes = {'fy': 0, 'sa': 0, 'acc': 0}
    for r in range(max(2, last_row - 60), last_row + 1):
        v = str(ws.cell(r, col).value or '').strip().upper()
        if v.startswith('SA'):
            votes['sa'] += 1
        elif v.startswith(('FY', 'YE')) or re.fullmatch(r'(19|20)\d{2}', v):
            votes['fy'] += 1
        elif _YEAR_LABEL.match(v):
            votes['acc'] += 1
    return max(votes, key=votes.get)


def _tokens(h: str) -> set:
    return set(h.split())


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith('=')


def _year_like_column(ws, col: int, last_row: int) -> bool:
    """Does this column hold fiscal-year labels (FY24, 23/24, 2025 ...)?"""
    vals = []
    for r in range(max(2, last_row - 60), last_row + 1):
        v = ws.cell(r, col).value
        if v not in (None, '') and not _is_formula(v):
            vals.append(str(v).strip())
    if len(vals) < 5:
        return False
    hits = sum(1 for v in vals if _YEAR_LABEL.match(v))
    # a year column repeats a handful of labels down hundreds of rows
    return hits >= 0.85 * len(vals) and len(set(vals)) <= max(10, len(vals) // 4)


def resolve_columns(ws, last_row: int = None) -> dict:
    """Return {role: column index or None, 'reviewed': [col, ...]}.

    last_row – last row holding client data (used to sample column contents);
    defaults to the sheet's last row."""
    hdrs = read_headers(ws)
    taken = set()

    def claim(col):
        if col:
            taken.add(col)
        return col

    cols = {
        'no':     named(hdrs, 'No', 'No.', 'Num', 'Seq'),
        'date':   named(hdrs, *DATE_HEADERS),
        'desc':   named(hdrs, *DESC_COLS),
        'type':   named(hdrs, 'Type', 'Transaction Type', 'Transaction type'),
        'in':     named(hdrs, *IN_COLS),
        'out':    named(hdrs, *OUT_COLS),
        'amt':    named(hdrs, 'Amount', 'Value'),
        'bal':    named(hdrs, 'Balance', 'Running Balance'),
        'net':    named(hdrs, 'NET', 'Net'),
        'buc':    named(hdrs, 'Balance UC'),
        'chk':    named(hdrs, 'Check UC', 'Check'),
        'ts':     named(hdrs, 'Timestamp'),
        'from':   named(hdrs, 'From'),
        'to':     named(hdrs, 'To'),
        'status': named(hdrs, 'Status'),
        'tag':    named(hdrs, 'Tag 1', 'Tag'),
        'sa':     named(hdrs, *_SA_NAMES),
        'acc':    named(hdrs, *_ACC_NAMES),
        'fy':     named(hdrs, *_FY_NAMES),
        'ref':    named(hdrs, *_REF_NAMES),
        'bank_cat': named(hdrs, *_BANK_CAT_NAMES),
        'uc':     named(hdrs, *UC_COLS),
        'correct': named(hdrs, *CORRECTION_COLS),
    }
    # A separate payee column next to the description (Monzo: 'Name' holds the
    # merchant, 'Description' the raw card text or a reference code).
    name_col = named(hdrs, 'Name', 'Payee', 'Payee Name', 'Merchant', 'Merchant Name',
                     'Counter Party', 'Counterparty')
    cols['name'] = name_col if name_col and name_col != cols['desc'] else None
    for c in cols.values():
        claim(c)

    # Column contents are sampled from the client's real rows, not from the
    # formatted-but-empty tail that openpyxl counts in max_row.
    if not last_row:
        last_row = last_data_row(ws, [c for c in (cols['date'], cols['desc']) if c] or [1])

    ordered = sorted(hdrs.items(), key=lambda kv: kv[1])     # left to right

    # ── category columns, by the words in the header ─────────────────────────
    if not cols['uc']:
        cands = [(c, h) for h, c in ordered
                 if c not in taken and 'uc' in _tokens(h)
                 and not _tokens(h) & {'balance', 'check', 'bal'}
                 and _tokens(h) & {'category', 'cat', 'desc', 'description', 'class'}]
        if cands:
            cols['uc'] = claim(cands[-1][0])
    # No 'uc'-worded column anywhere: a plain 'Category' (no bank/spending/
    # statement qualifier, so it isn't the bank's own import category) is
    # where the client wants the AI's answer written.
    if not cols['uc']:
        cands = [(c, h) for h, c in ordered
                 if c not in taken
                 and _tokens(h) & {'category', 'cat', 'class'}
                 and not _tokens(h) & {'balance', 'check', 'bal', 'bank', 'spending', 'statement'}]
        if cands:
            cols['uc'] = claim(cands[-1][0])
    if not cols['bank_cat']:
        cands = [(c, h) for h, c in ordered
                 if c not in taken and 'categor' in h
                 and _tokens(h) & {'bank', 'spending', 'statement'}]
        if not cands and cols['uc']:
            # a plain 'Category' next to a UC column is the bank's own category
            cands = [(c, h) for h, c in ordered if c not in taken and h == 'category']
        if cands:
            cols['bank_cat'] = claim(cands[0][0])

    # A second category column next to 'uc': the reviewer's FINAL answer.
    # The AI's first pass goes in 'uc'; a senior reviewer writes the final
    # category here, and the client's pivot / Analysis are built on it. New
    # rows leave it empty for the reviewer; it is the column the AI learns
    # from. Two naming habits produce this layout — 'MUC desc' next to
    # 'UC desc' ('muc' token), or a second 'UC Category 1' next to
    # 'UC Category' (same tokens, a trailing number) — so after the 'muc'
    # case, a column is also recognised this way when it carries the same
    # 'uc' wording as the chosen column but is far more completely filled:
    # the draft column is the one with the gaps, since not every row has
    # been reviewed yet.
    def _fill_rate(col):
        vals = [ws.cell(r, col).value for r in range(2, last_row + 1)]
        vals = [v for v in vals if v not in (None, '') and not _is_formula(v)]
        return len(vals) / max(1, last_row - 1)

    cols['reviewed'] = [c for h, c in ordered
                        if c not in taken and 'muc' in _tokens(h) and cols['uc']]
    if not cols['reviewed'] and cols['uc']:
        uc_fill = _fill_rate(cols['uc'])
        cols['reviewed'] = [c for h, c in ordered
                            if c not in taken and 'uc' in _tokens(h)
                            and _tokens(h) & {'category', 'cat', 'desc', 'description', 'class'}
                            and _fill_rate(c) >= 0.9 and _fill_rate(c) - uc_fill >= 0.2]
    for c in cols['reviewed']:
        claim(c)

    # ── year columns, by alias then by contents ─────────────────────────────
    # Year columns with unfamiliar names ('Year', 'Personal') are recognised
    # by their contents and filed by the style of label they hold, so the
    # column holding FY25 is the 'fy' one whatever it is called.
    free_slots = [k for k in ('acc', 'fy', 'sa') if not cols[k]]
    if free_slots:
        for h, c in ordered:
            if not free_slots:
                break
            if c not in taken and c != cols['date'] and _year_like_column(ws, c, last_row):
                style = _year_style(ws, c, last_row)
                slot = style if style in free_slots else free_slots[0]
                free_slots.remove(slot)
                cols[slot] = claim(c)

    return cols
