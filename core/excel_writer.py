"""Dynamic Excel writer — appends transactions to any client template while
keeping the output identical in style to the input template.

Template preservation:
 - New rows copy every cell's style (font, borders, fill, number format)
   from the client's last existing data row.
 - Formula columns are replicated from the template's own formulas
   (row references translated), so NET / Balance UC / Check behave exactly
   like the client's existing rows.
 - Date cells match the template's type: text dates stay text in the same
   format; real dates keep the template's number format.
 - Year labels (SA / FY / ACC) are LEARNED from the client's existing rows
   (see core/year_labels.py) — every client's fiscal-year convention is
   reproduced, whatever it is.
 - The Analysis sheet is updated in place (values only) instead of being
   deleted and rebuilt, so its formatting survives.
"""

import re
import shutil
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.properties import CalcProperties

from core.financial_year import get_fy
from parsers._common import norm_header
from core.year_labels import YearLabeller, learn_year_columns


# ── Sheet / header helpers ────────────────────────────────────────────────────

def _detect_sheet(wb: openpyxl.Workbook, keyword: str) -> str:
    """Find the RAW transactions sheet or the Analysis summary sheet.

    First tries the sheet name — most templates have 'raw' and 'analysis' in
    the title. When nothing matches (clients use 'Transactions', 'Summary',
    their trading name, etc.), scores every sheet by content and picks the
    best match. The transactions sheet looks like: many rows, a Date header
    and money columns in row 1. The analysis sheet looks like: fewer rows,
    formulas pointing elsewhere, and no transaction headers."""
    kw = keyword.lower()
    matches = [name for name in wb.sheetnames if kw in name.lower()]
    if kw == 'raw':
        pure = [m for m in matches if 'analysis' not in m.lower()
                and 'summary' not in m.lower()]
        if pure:
            return pure[0]
    if matches:
        return matches[0]

    # Content-based fallback — scores each sheet by how much it looks like
    # the kind we're after.
    best, best_score = None, -1
    for name in wb.sheetnames:
        ws = wb[name]
        score = _sheet_score(ws, kw)
        if score > best_score:
            best, best_score = name, score
    if best is None or best_score <= 0:
        raise KeyError(f"No {keyword} sheet found in {wb.sheetnames} "
                       f"(looked by name and by content)")
    return best


def _sheet_score(ws, kind: str) -> int:
    """How likely is this sheet to be a 'raw' transactions sheet or an
    'analysis' summary? Higher is better. Transactions: has Date + a money
    column header in row 1, and many rows. Analysis: has formulas and few
    rows."""
    headers = [str(c.value or '').strip().lower()
               for c in next(ws.iter_rows(min_row=1, max_row=1), ())]
    row_count = ws.max_row
    formula_count = sum(
        1 for row in ws.iter_rows(min_row=1, max_row=min(row_count, 100))
        for c in row if isinstance(c.value, str) and c.value.startswith('=')
    )
    has_date = any(h in {'date', 'transaction date', 'posting date'} for h in headers)
    has_money = any(h in {'amount', 'net', 'money in', 'money out', 'paid in',
                          'paid out', 'credit', 'debit', 'in', 'out', 'value',
                          'balance'} for h in headers)
    has_uc = any('category' in h or h == 'uc' for h in headers)

    if kind == 'raw':
        if not has_date or not has_money:
            return 0
        return row_count + (50 if has_uc else 0)
    # kind == 'analysis'
    if has_date and has_money and row_count > 50:
        return 0        # looks like a transactions sheet, not a summary
    return formula_count + (20 if row_count < 100 else 0)


def _read_headers(ws) -> dict:
    row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    hdrs = {}
    for i, v in enumerate(row1):
        if v is not None and str(v).strip():
            hdrs.setdefault(norm_header(v), i + 1)
    return hdrs


# Template column names, matched ignoring case / spacing / '(£)' / '(GBP)'.
DESC_COLS = ('Counter Party', 'Counterparty', 'Description', 'Transaction description',
             'Transaction Details', 'Details', 'Memo', 'Narrative', 'Particulars',
             'Payee', 'Name', 'Merchant')
IN_COLS   = ('Paid in', 'Money in', 'In', 'Credit', 'Credits', 'Receipts', 'Deposits')
OUT_COLS  = ('Paid out', 'Money out', 'Out', 'Debit', 'Debits', 'Payments',
             'Withdrawals', 'Withdrawn')
UC_COLS   = ('UC category', 'UC Category', 'UC')
# Where the accountant types the right category when the AI's is wrong
# (read by scripts/learn_from_feedback.py).
CORRECTION_COLS = ('Correct category', 'Correction', 'Accountant category')
CORRECTION_HEADER = CORRECTION_COLS[0]

# Rows the AI was unsure about: highlighted on the category cell.
REVIEW_FILL = PatternFill('solid', start_color='FFFF00', end_color='FFFF00')
REVIEW_NOTE = ('Please check: the AI was not sure about this category. If it is '
               f"wrong, type the right one in the '{CORRECTION_HEADER}' column.")


def _col(hdrs: dict, *names: str):
    for n in names:
        idx = hdrs.get(norm_header(n))
        if idx is not None:
            return idx
    return None


def _find_last_data_row(ws, key_cols: list) -> int:
    """Last row where any key column has a value."""
    last = 1
    for r in range(2, ws.max_row + 1):
        for c in key_cols:
            if c and ws.cell(r, c).value not in (None, ''):
                last = r
                break
    return last


# ── PivotTable sync ────────────────────────────────────────────────────────

_REF_RE = re.compile(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$')


def _relocate_pivot_clear_of_grid(ws_a, grid_last_col: int) -> None:
    """Move any PivotTable anchored inside the Analysis sheet's own SUMIFS
    grid (columns 1..grid_last_col) out to empty columns on its right. The
    grid (columns 1..grid_last_col) out to empty columns on its right. The
    grid and the pivot used to share the same cells — refreshing the pivot
    in place made Excel treat the grid's values as "data in the way" and
    prompt to overwrite them. The target columns must be empty over the
    pivot's rows (plus room to grow for new years/categories), so the pivot
    never lands on the accountant's own workings either. Idempotent: a
    pivot already clear of the grid is left alone."""
    occupied = {}
    for row in ws_a.iter_rows():
        for cell in row:
            if cell.value not in (None, ''):
                occupied.setdefault(cell.column, []).append(cell.row)
    for piv in getattr(ws_a, '_pivots', []):
        loc = piv.location
        m = _REF_RE.match(loc.ref or '')
        if not m:
            continue
        c1, r1, c2, r2 = m.groups()
        start_col = column_index_from_string(c1)
        if start_col > grid_last_col + 1:
            continue
        end_col = column_index_from_string(c2)
        width = end_col - start_col + 1
        top, bottom = int(r1), int(r2) + 20          # room to grow downwards

        def free(c):
            return not any(top <= r <= bottom for r in occupied.get(c, []))

        target = grid_last_col + 2
        # need width + 2 spare columns (growth to the right) plus a gap column
        while not all(free(c) for c in range(target - 1, target + width + 2)):
            target += 1
        delta = target - start_col
        loc.ref = (f'{get_column_letter(start_col + delta)}{r1}:'
                   f'{get_column_letter(end_col + delta)}{r2}')
        _repoint_getpivotdata(ws_a, c1, r1, get_column_letter(start_col + delta))


def _repoint_getpivotdata(ws_a, old_col: str, row: str, new_col: str) -> None:
    """The accountant's GETPIVOTDATA(..., $A$3, ...) formulas name the pivot
    by a cell inside it; once the pivot moves, point them at its new place
    or they all turn into #REF!. Covers the pivot's own sheet (plain $A$3)
    and every other sheet ('Analysis 25'!$A$3 / Analysis!$A$3)."""
    title = ws_a.title
    own = r'(?<![!A-Za-z0-9_$])'
    other = "(?:'" + re.escape(title.replace("'", "''")) + "'|" + re.escape(title) + ")!"
    cell = r'(\$?)' + old_col + r'(\$?)' + row + r'(?!\d)'
    head = r'(GETPIVOTDATA\(\s*"[^"]*"\s*,\s*)'
    rx_own = re.compile(head + own + cell, re.I)
    rx_other = re.compile(head + '(' + other + ')' + cell, re.I)
    for ws in ws_a.parent.worksheets:
        for line in ws.iter_rows():
            for c in line:
                v = c.value
                if not (isinstance(v, str) and v.startswith('=') and 'GETPIVOTDATA' in v.upper()):
                    continue
                if ws is ws_a:
                    v = rx_own.sub(lambda m: f'{m[1]}{m[2]}{new_col}{m[3]}{row}', v)
                v = rx_other.sub(lambda m: f'{m[1]}{m[2]}{m[3]}{new_col}{m[4]}{row}', v)
                c.value = v


def _align_pivot_columns(wb, sheet: str, cache, col1: str, row1: str, col2: str):
    """A pivot whose source range is shifted off its own columns (e.g. B:J
    while its fields are the headers in A:I) fails on refresh — Excel finds a
    blank or wrong header. Return the columns whose headers match the pivot's
    fields, or the original columns when they already match or no match is
    found."""
    names = [str(f.name or '').strip() for f in (cache.cacheFields or [])]
    if not names or sheet not in wb.sheetnames:
        return col1, col2
    ws = wb[sheet]
    hr = int(row1)

    def headers_at(start):
        return [str(ws.cell(hr, start + k).value or '').strip() for k in range(len(names))]

    c1 = column_index_from_string(col1)
    if headers_at(c1) == names:
        return col1, col2
    for start in range(1, ws.max_column - len(names) + 2):
        if headers_at(start) == names:
            return get_column_letter(start), get_column_letter(start + len(names) - 1)
    return col1, col2


def _refresh_pivot_sources(wb, new_last_row: int, raw_sheet: str = None) -> None:
    """Grow every embedded PivotTable's source range to cover the newly
    appended rows and mark it to refresh on open, so a new fiscal year
    (e.g. FY26) shows up inside the pivot automatically — no manual
    'Refresh' click needed. Safe to call only after any pivot sharing a
    sheet with an Analysis grid has been relocated clear of it (see
    _relocate_pivot_clear_of_grid); otherwise Excel prompts to overwrite
    the grid's values on every open."""
    for sheet in wb.worksheets:
        for piv in getattr(sheet, '_pivots', []):
            cache = piv.cache
            src = cache.cacheSource
            if src is None or src.type != 'worksheet' or src.worksheetSource is None:
                continue
            wsrc = src.worksheetSource
            # A pivot whose source sheet was renamed (e.g. 'RAW 2526' when
            # the tab is now 'RAW 25') fails on refresh — repoint it.
            if raw_sheet and wsrc.sheet and wsrc.sheet not in wb.sheetnames:
                wsrc.sheet = raw_sheet
            m = _REF_RE.match(wsrc.ref or '')
            if not m:
                continue
            col1, row1, col2, row2 = m.groups()
            col1, col2 = _align_pivot_columns(wb, wsrc.sheet, cache, col1, row1, col2)
            end_row = max(int(row2), new_last_row)
            wsrc.ref = f'{col1}{row1}:{col2}{end_row}'
            cache.refreshOnLoad = True


# ── Date handling ─────────────────────────────────────────────────────────────

_DATE_FMTS = ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d %b %Y', '%d/%m/%y', '%m/%d/%Y')


def _detect_text_date_format(sample: str):
    s = str(sample).strip()
    for fmt in _DATE_FMTS:
        try:
            datetime.strptime(s, fmt)
            return fmt
        except ValueError:
            continue
    return None


# ── Formula translation ───────────────────────────────────────────────────────

_CELL_REF = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def _translate_formula(formula: str, src_row: int, dst_row: int) -> str:
    """Shift relative row references near src_row to dst_row.

    Absolute row refs ($4) and far-away refs (headers etc.) are kept."""
    delta = dst_row - src_row

    def repl(m):
        col_abs, col, row_abs, row_s = m.groups()
        row = int(row_s)
        if row_abs == '$':
            return m.group()
        if abs(row - src_row) <= 1:
            return f'{col_abs}{col}{row_abs}{row + delta}'
        return m.group()

    return _CELL_REF.sub(repl, formula)


# ── Client example extraction (teaches the AI this client's vocabulary) ──────

def extract_client_examples(template_path, sheet_name: str = None) -> list:
    """Return [{'description','type','reference','category'}] from the
    template's existing rows — the accountant's own past decisions."""
    wb = openpyxl.load_workbook(template_path, data_only=True)
    try:
        raw_name = sheet_name if (sheet_name and sheet_name in wb.sheetnames) \
            else _detect_sheet(wb, 'raw')
    except KeyError:
        return []
    ws = wb[raw_name]
    hdrs = _read_headers(ws)

    desc_col = _col(hdrs, *DESC_COLS)
    uc_col   = _col(hdrs, *UC_COLS)
    ref_col  = _col(hdrs, 'Reference', 'Ref')
    type_col = _col(hdrs, 'Type', 'Transaction Type', 'Transaction type')
    in_col   = _col(hdrs, *IN_COLS)
    out_col  = _col(hdrs, *OUT_COLS)
    amt_col  = _col(hdrs, 'Amount', 'Value', 'NET', 'Net')

    if not desc_col or not uc_col:
        return []

    def _num(c, r):
        v = ws.cell(r, c).value if c else None
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0

    def direction(r):
        """'in' / 'out' for the row's money, None when it can't be told."""
        if in_col and _num(in_col, r) > 0:
            return 'in'
        if out_col and _num(out_col, r):
            return 'out'
        amt = _num(amt_col, r)
        return 'in' if amt > 0 else 'out' if amt < 0 else None

    examples = []
    for r in range(2, ws.max_row + 1):
        desc = ws.cell(r, desc_col).value
        cat  = ws.cell(r, uc_col).value
        if not desc or not cat:
            continue
        cat = str(cat).strip()
        if not cat or cat.startswith('='):
            continue
        examples.append({
            'description': str(desc).strip(),
            'type':      str(ws.cell(r, type_col).value or '').strip() if type_col else '',
            'reference': str(ws.cell(r, ref_col).value  or '').strip() if ref_col  else '',
            'category':  cat,
            'direction': direction(r),
        })
    return examples


# A criteria range on this sheet's column A (A:A or $A$5:$A$40) followed by
# the criteria cell — covers SUMIF(A:A,K22,D:D) and SUMIFS(D:D,A:A,K22).
_SUMIF_A_RE = re.compile(
    r"(?<![A-Za-z0-9_!$'])\$?A(?:\$?\d+)?:\$?A(?:\$?\d+)?\s*,\s*\$?([A-Z]{1,3})\$?(\d+)(?![\d(])")


def extract_analysis_vocabulary(template_path) -> list:
    """Category names the accountant's own Analysis workings look up, e.g.
    'Inusrance - AOP' / 'Subscription' / 'Charity' in a P&L block built from
    =SUMIF(A:A,K22,D:D). New rows must use exactly these names, or those
    formulas never pick them up. Returns the labels in sheet order."""
    wb = openpyxl.load_workbook(template_path)
    try:
        ws = wb[_detect_sheet(wb, 'analysis')]
    except KeyError:
        return []
    out = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith('=') and 'SUMIF' in v.upper()):
                continue
            for col, r in _SUMIF_A_RE.findall(v):
                label = ws[f'{col}{r}'].value
                # 'Equipment?' is the accountant's own query, not a category
                if isinstance(label, str) and label.strip() and not label.startswith('=') \
                        and not label.strip().endswith('?') and label not in out:
                    out.append(label)
    return out


# ── New transactions only, with complete balances ─────────────────────────────

def _signed(t: dict) -> float:
    return round((t.get('money_in') or 0) - (t.get('money_out') or 0), 2)


def fill_missing_balances(transactions: list) -> int:
    """Some statements (e.g. Starling PDFs) print a balance only on the last
    transaction of each day. Fill the gaps from the neighbouring printed
    balances, but only where the run of filled rows lands exactly on the next
    printed balance — so a filled balance is always one the bank's own
    figures prove. Returns how many balances were filled."""
    n = len(transactions)
    filled = 0

    # Gaps at the very start or end have a printed balance on one side only,
    # so they can't be checked on their own. Fill them only when the whole
    # statement is proven to run in this order: every printed balance equals
    # the one before it plus the amounts in between (a newest-first
    # statement, or a mis-parsed amount, fails this and they stay blank).
    known = [k for k, t in enumerate(transactions) if t.get('balance') is not None]
    order_proven = len(known) >= 2 and all(
        abs(round(transactions[a]['balance']
                  + sum(_signed(t) for t in transactions[a + 1:b + 1]), 2)
            - transactions[b]['balance']) < 0.005
        for a, b in zip(known, known[1:]))

    i = 0
    while i < n:
        if transactions[i].get('balance') is not None:
            i += 1
            continue
        j = i                          # transactions[i:j] have no balance
        while j < n and transactions[j].get('balance') is None:
            j += 1
        prev = transactions[i - 1]['balance'] if i > 0 else None
        nxt = transactions[j]['balance'] if j < n else None
        run = []
        if prev is not None:
            b = prev
            for t in transactions[i:j]:
                b = round(b + _signed(t), 2)
                run.append(b)
            if nxt is None:            # gap at the very end
                ok = order_proven
            else:
                ok = abs(b + _signed(transactions[j]) - nxt) < 0.005
        elif nxt is not None:          # gap at the very start: work backwards
            b = nxt
            for t in reversed(transactions[i + 1:j + 1]):
                b = round(b - _signed(t), 2)
                run.append(b)
            run.reverse()
            ok = order_proven
        else:
            ok = False
        if ok:
            for t, b in zip(transactions[i:j], run):
                t['balance'] = b
            filled += j - i
        i = j
    return filled


def drop_rows_already_in_template(transactions: list, template_path,
                                  sheet_name: str = None) -> tuple[list, int]:
    """Remove transactions the template's RAW sheet already holds — a new
    statement usually overlaps the last few days already entered. A row
    matches when date and amount agree and so does either the balance or
    the counterparty; each existing row can match only one new transaction,
    so genuine same-day repeats are kept. Returns (kept, dropped count)."""
    wb = openpyxl.load_workbook(template_path, data_only=True)
    try:
        ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) \
            else wb[_detect_sheet(wb, 'raw')]
    except KeyError:
        return transactions, 0
    hdrs = _read_headers(ws)
    date_col = _col(hdrs, 'Date')
    desc_col = _col(hdrs, *DESC_COLS)
    in_col, out_col = _col(hdrs, *IN_COLS), _col(hdrs, *OUT_COLS)
    amt_col = _col(hdrs, 'Amount', 'Value', 'NET', 'Net')
    bal_col = _col(hdrs, 'Balance', 'Running Balance')
    if not date_col:
        return transactions, 0

    def num(r, c):
        v = ws.cell(r, c).value if c else None
        return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    def as_date(v):
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        if isinstance(v, str):
            for fmt in _DATE_FMTS:
                try:
                    return datetime.strptime(v.strip(), fmt).date()
                except ValueError:
                    pass
        return None

    def desc_key(s):
        return re.sub(r'[^a-z0-9]+', '', str(s or '').lower())

    existing = {}      # (date, amount) -> [[balance, desc key, used], ...]
    for r in range(2, ws.max_row + 1):
        d = as_date(ws.cell(r, date_col).value)
        if d is None:
            continue
        amt = num(r, amt_col)
        if amt is None and (in_col or out_col):
            p_in, p_out = num(r, in_col) or 0, num(r, out_col) or 0
            amt = p_in - abs(p_out) if (p_in or p_out) else None
        if amt is None:
            continue
        existing.setdefault((d, round(amt, 2)), []).append(
            [num(r, bal_col), desc_key(ws.cell(r, desc_col).value if desc_col else ''), False])

    kept, dropped = [], 0
    for t in transactions:
        d = t['date'].date() if isinstance(t['date'], datetime) else t['date']
        bal, dk = t.get('balance'), desc_key(t.get('description'))
        match = None
        for row in existing.get((d, _signed(t)), []):
            if row[2]:
                continue
            same_bal = bal is not None and row[0] is not None and abs(bal - row[0]) < 0.005
            if same_bal or (dk and dk == row[1]):
                match = row
                break
        if match:
            match[2] = True
            dropped += 1
            print(f"  [excel_writer]   already in template, skipped: {d} "
                  f"{_signed(t):+.2f} {t.get('description') or ''}")
        else:
            kept.append(t)
    return kept, dropped


def prepare_new_transactions(transactions: list, template_path,
                             sheet_name: str = None) -> tuple[list, int]:
    """Fill statement balance gaps, then drop rows already in the template.
    Returns (transactions to add, number dropped as already present)."""
    filled = fill_missing_balances(transactions)
    if filled:
        print(f"  [excel_writer] Filled {filled} balance(s) the statement left blank "
              f"(checked against its printed balances)")
    kept, dropped = drop_rows_already_in_template(transactions, template_path, sheet_name)
    if dropped:
        print(f"  [excel_writer] Skipped {dropped} transaction(s) already in the template")
    return kept, dropped


# ── Fallback year formatting (only when template has no existing rows) ───────

def _fallback_year_label(d, style: str):
    sa, acc = get_fy(d)          # sa='FY24' start-year, acc='24/25'
    end = acc[3:]
    if style == 'sa':
        return f"SA{acc}"        # 'SA24/25'
    if style == 'acc':
        return acc               # '24/25'
    if style == 'fy':
        return f"FY{end}"        # END-year label: FY25 = year 2024/25
    return None


# ── Main entry point ──────────────────────────────────────────────────────────

def write_workbook(
    transactions: list,
    categories: list,
    template_path,
    output_path,
    sheet_name: str = None,
    review: list = None,
) -> None:
    """review – optional list of bools, one per transaction: True highlights
    that row's category for the accountant to check."""
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)

    # ── RAW sheet ─────────────────────────────────────────────────────────────
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb[_detect_sheet(wb, 'raw')]

    hdrs = _read_headers(ws)

    no_col    = _col(hdrs, 'No', '#')
    sa_col    = _col(hdrs, 'SA')
    acc_col   = _col(hdrs, 'ACC', 'AC')
    fy_col    = _col(hdrs, 'FY')
    date_col  = _col(hdrs, 'Date')
    desc_col  = _col(hdrs, *DESC_COLS)
    ref_col   = _col(hdrs, 'Reference', 'Ref')
    type_col  = _col(hdrs, 'Type', 'Transaction Type', 'Transaction type')
    in_col    = _col(hdrs, *IN_COLS)
    out_col   = _col(hdrs, *OUT_COLS)
    amt_col   = _col(hdrs, 'Amount', 'Value')
    bal_col   = _col(hdrs, 'Balance', 'Running Balance')
    net_col   = _col(hdrs, 'NET', 'Net')
    buc_col   = _col(hdrs, 'Balance UC')
    chk_col   = _col(hdrs, 'Check UC', 'Check')
    csv_cat_col = _col(hdrs, 'Spending Category', 'Category name')
    uc_col    = _col(hdrs, *UC_COLS)
    ts_col    = _col(hdrs, 'Timestamp')
    from_col   = _col(hdrs, 'From')
    to_col     = _col(hdrs, 'To')
    status_col = _col(hdrs, 'Status')
    tag_col    = _col(hdrs, 'Tag 1', 'Tag')

    # Templates with no Paid in / Paid out and no Amount column keep the
    # signed amount straight in 'Net' as a plain value — write it there.
    # A Net column holding formulas is left to the formula replication below.
    if not amt_col and not (in_col and out_col) and net_col:
        net_has_formula = any(
            isinstance(ws.cell(r, net_col).value, str)
            and ws.cell(r, net_col).value.startswith('=')
            for r in range(2, ws.max_row + 1)
        )
        if not net_has_formula:
            amt_col = net_col

    # Say loudly when the template can't hold key transaction data — the
    # rows would otherwise go out with those fields silently missing.
    missing = [label for label, ok in (
        ('Date', date_col), ('Description', desc_col),
        ('Paid in / Paid out or Amount', (in_col and out_col) or amt_col),
    ) if not ok]
    if missing:
        print(f"  [excel_writer] WARNING: template sheet '{ws.title}' has no "
              f"{', '.join(missing)} column — headers are: {list(hdrs)}")

    # A real transaction row has a date or description. Some templates carry
    # hundreds of pre-numbered empty rows (only 'No' filled): counting those
    # as data made new rows land below them — leaving a '(blank)' year in the
    # pivot — and made the empty row the style/formula model, so formulas
    # such as NET were never copied onto the new rows.
    content_cols = [c for c in (date_col, desc_col) if c] or [no_col]
    last_row = _find_last_data_row(ws, content_cols)
    numbered_last = _find_last_data_row(ws, [no_col]) if no_col else last_row

    # ── Learn this client's year-label conventions from their own rows ───────
    labellers = learn_year_columns(
        ws, {'sa': sa_col, 'acc': acc_col, 'fy': fy_col}, date_col, last_row,
    )

    # ── Template rows used as style/formula models for appended rows ─────────
    model_row = last_row if last_row > 1 else None
    max_col = ws.max_column

    model_styles   = {}
    model_formulas = {}   # col -> (formula, source_row)
    date_text_fmt  = None
    if model_row:
        for c in range(1, max_col + 1):
            cell = ws.cell(model_row, c)
            model_styles[c] = copy(cell._style)
        # Collect the most recent formula per column, scanning several rows up:
        # some templates carry a formula only on the row-sign it applies to
        # (e.g. IN = IF(H>0,H,"") appears only on income rows).
        scan_from = max(2, model_row - 200)
        for r_scan in range(model_row, scan_from - 1, -1):
            for c in range(1, max_col + 1):
                if c in model_formulas:
                    continue
                v = ws.cell(r_scan, c).value
                if isinstance(v, str) and v.startswith('='):
                    model_formulas[c] = (v, r_scan)
        d_model = ws.cell(model_row, date_col).value if date_col else None
        if isinstance(d_model, str):
            date_text_fmt = _detect_text_date_format(d_model)

    # Sign-conditional IN/OUT formulas: when BOTH sides carry a formula in the
    # template, each appended row gets only the side matching its sign — the
    # other cell stays empty, exactly like the client's own rows.
    sign_conditional = (
        in_col in model_formulas and out_col in model_formulas
        and amt_col is not None
    )

    # Paid-out sign convention: some templates (e.g. Monzo layouts) keep
    # money out as a NEGATIVE number with NET = in + out; others keep it
    # positive with NET = in - out. Follow whatever the client's rows do.
    out_negative = False
    if out_col and out_col not in model_formulas:
        neg = pos = 0
        for r_scan in range(2, last_row + 1):
            v = ws.cell(r_scan, out_col).value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and v:
                neg += v < 0
                pos += v > 0
        out_negative = neg > pos

    # ── Spending Category learned from the client's own rows ─────────────────
    # PDF statements carry no spending category, so new rows would come out
    # blank. Reuse what the client's history says: the category this counter
    # party had before, else the usual one for the row's UC category.
    def _key(s):
        return re.sub(r'[^a-z0-9]+', ' ', str(s or '').lower()).strip()

    spend_by_desc, spend_by_uc = {}, {}
    if csv_cat_col:
        for r_scan in range(2, last_row + 1):
            sc = ws.cell(r_scan, csv_cat_col).value
            if not isinstance(sc, str) or not sc.strip() or sc.startswith('='):
                continue
            sc = sc.strip()
            if desc_col:
                dk = _key(ws.cell(r_scan, desc_col).value)
                if dk:
                    spend_by_desc.setdefault(dk, {}).setdefault(sc, 0)
                    spend_by_desc[dk][sc] += 1
            if uc_col:
                uk = _key(ws.cell(r_scan, uc_col).value)
                if uk:
                    spend_by_uc.setdefault(uk, {}).setdefault(sc, 0)
                    spend_by_uc[uk][sc] += 1

    def learned_spending(desc, cat):
        votes = spend_by_desc.get(_key(desc)) or spend_by_uc.get(_key(cat))
        return max(votes, key=votes.get) if votes else None

    # ── Other template columns filled from the statement's own columns ───────
    # e.g. 'Account Name' / 'Account Number': when the statement has a column
    # with the same header, its value is copied across. Text stays text (an
    # account number like '600260-10322221' must not become a number); plain
    # numbers become numbers when the template's own rows hold numbers there.
    mapped = {c for c in (no_col, sa_col, acc_col, fy_col, date_col, desc_col,
                          ref_col, type_col, in_col, out_col, amt_col, bal_col,
                          net_col, buc_col, chk_col, csv_cat_col, uc_col, ts_col,
                          from_col, to_col, status_col, tag_col) if c}
    passthrough = {}    # normalised header -> template column
    for h, c in hdrs.items():
        if c not in mapped and c not in model_formulas and not _col({h: c}, *CORRECTION_COLS):
            passthrough[h] = c
    numeric_cols = set()
    if model_row:
        for c in passthrough.values():
            v = ws.cell(model_row, c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                numeric_cols.add(c)

    def passthrough_value(col, v):
        if isinstance(v, str):
            v = v.strip()
            if col in numeric_cols and re.fullmatch(r'-?\d[\d,]*(\.\d+)?', v):
                return float(v.replace(',', ''))
        return v if v not in (None, '') else None

    # ── Highest existing sequence number ──────────────────────────────────────
    last_no = 0
    if no_col:
        for r in range(2, last_row + 1):
            v = ws.cell(r, no_col).value
            try:
                last_no = max(last_no, int(v))
            except (TypeError, ValueError):
                pass

    def year_label(kind: str, fallback_style: str, d):
        lab: YearLabeller = labellers.get(kind)
        if lab:
            return lab.label_for(d)
        return _fallback_year_label(d, fallback_style)

    # Categories go in with the client's exact spelling — 'Professional fee '
    # with its trailing space — since SUMIFS / VLOOKUP match text exactly.
    raw_spelling = {}
    if uc_col:
        for r_scan in range(2, last_row + 1):
            v = ws.cell(r_scan, uc_col).value
            if isinstance(v, str) and v.strip() and not v.startswith('='):
                raw_spelling.setdefault(v.strip(), v)

    # ── Write transaction rows ────────────────────────────────────────────────
    for i, (txn, cat) in enumerate(zip(transactions, categories)):
        r = last_row + 1 + i
        seq = last_no + 1 + i
        if isinstance(cat, str):
            cat = raw_spelling.get(cat.strip(), cat)

        d = txn['date']
        if isinstance(d, datetime):
            d = d.date()

        money_in  = txn.get('money_in', 0.0)
        money_out = txn.get('money_out', 0.0)
        balance   = txn.get('balance')
        desc      = (txn.get('description') or '').strip() or (txn.get('subcategory') or '').strip()

        # 1. apply the template's styles to the whole new row first
        for c in range(1, max_col + 1):
            if c in model_styles:
                ws.cell(r, c)._style = copy(model_styles[c])

        def w(col, val):
            if col and val is not None and col not in model_formulas:
                ws.cell(r, col, value=val)

        # 2. data values
        w(no_col, seq)
        if sa_col:
            w(sa_col, year_label('sa', 'sa', d))
        if acc_col:
            w(acc_col, year_label('acc', 'acc', d))
        if fy_col:
            w(fy_col, year_label('fy', 'fy', d))

        if date_col and date_col not in model_formulas:
            if date_text_fmt:
                ws.cell(r, date_col, value=datetime(d.year, d.month, d.day).strftime(date_text_fmt))
            else:
                ws.cell(r, date_col, value=datetime(d.year, d.month, d.day))

        w(desc_col, desc or None)
        w(ref_col,  (txn.get('reference') or '').strip() or None)
        w(type_col, (txn.get('subcategory') or '').strip() or None)
        w(in_col,   money_in  if money_in  > 0 else None)
        w(out_col,  (-money_out if out_negative else money_out) if money_out > 0 else None)
        w(bal_col,  balance)
        w(csv_cat_col, (txn.get('csv_category') or '').strip()
                       or learned_spending(desc, cat))
        w(uc_col,   cat)
        if uc_col and review and i < len(review) and review[i]:
            ws.cell(r, uc_col).fill = REVIEW_FILL
            ws.cell(r, uc_col).comment = Comment(REVIEW_NOTE, 'Plutus')
        w(ts_col,     txn.get('timestamp'))
        w(from_col,   (txn.get('from')   or '').strip() or None)
        w(to_col,     (txn.get('to')     or '').strip() or None)
        w(status_col, (txn.get('status') or '').strip() or None)
        w(tag_col,    (txn.get('tag')    or '').strip() or None)
        for h, v in (txn.get('extra') or {}).items():
            col = passthrough.get(norm_header(h))
            if col:
                w(col, passthrough_value(col, v))

        # signed amount column
        if amt_col and amt_col not in model_formulas:
            signed = money_in if money_in > 0 else (-money_out if money_out > 0 else None)
            if signed is not None:
                ws.cell(r, amt_col, value=signed)

        # 3. replicate the template's own formulas (NET, Balance UC, Check, …)
        for c, (f, src_row) in model_formulas.items():
            if sign_conditional:
                if c == in_col and money_in <= 0:
                    continue    # leave IN empty on money-out rows
                if c == out_col and money_out <= 0:
                    continue    # leave OUT empty on money-in rows
            ws.cell(r, c, value=_translate_formula(f, src_row, r))

        # 4. fallback formulas when the template rows carry static values
        if net_col and net_col not in model_formulas and in_col and out_col:
            in_l, out_l = get_column_letter(in_col), get_column_letter(out_col)
            op = '+' if out_negative else '-'
            ws.cell(r, net_col, value=f'={in_l}{r}{op}{out_l}{r}')
        if buc_col and buc_col not in model_formulas and net_col:
            buc_l, net_l = get_column_letter(buc_col), get_column_letter(net_col)
            if r - 1 >= 2:
                ws.cell(r, buc_col, value=f'={buc_l}{r - 1}+{net_l}{r}')
            else:
                ws.cell(r, buc_col, value=f'={net_l}{r}')
        if chk_col and chk_col not in model_formulas and bal_col and buc_col and balance is not None:
            bal_l, buc_l = get_column_letter(bal_col), get_column_letter(buc_col)
            ws.cell(r, chk_col, value=f'={bal_l}{r}-{buc_l}{r}')

    new_last = last_row + len(transactions)

    # The template's filter (e.g. A1:P103) must cover the new rows too, or
    # filtering in Excel silently leaves them out.
    m = _REF_RE.match(ws.auto_filter.ref or '')
    if m and int(m.group(4)) >= last_row:
        c1, r1, c2, _ = m.groups()
        ws.auto_filter.ref = f'{c1}{r1}:{c2}{max(new_last, int(m.group(4)))}'

    # A ready-made column for the accountant's corrections, right of the AI's
    # category when that column is completely free — otherwise (e.g. Ahmed's
    # 'Type' VLOOKUP sits there) the first empty column after the sheet's
    # data. learn_from_feedback finds it by its header, wherever it is.
    if uc_col and not _col(hdrs, *CORRECTION_COLS):
        def empty(c):
            return all(ws.cell(r, c).value in (None, '') for r in range(1, ws.max_row + 1))
        if empty(uc_col + 1):
            corr_col = uc_col + 1
        else:       # formatted-but-blank columns don't count as used
            corr_col = 1 + max(c for c in range(1, ws.max_column + 1) if not empty(c))
        if empty(corr_col):
            head = ws.cell(1, corr_col, value=CORRECTION_HEADER)
            head._style = copy(ws.cell(1, uc_col)._style)
            ws.column_dimensions[get_column_letter(corr_col)].width = max(
                ws.column_dimensions[get_column_letter(uc_col)].width or 0, 18)

    # Pre-numbered empty rows left over below the new data would show up as
    # a '(blank)' year in the pivot and carry duplicate numbers — clear them.
    if no_col:
        for r in range(new_last + 1, numbered_last + 1):
            if all(ws.cell(r, c).value in (None, '') for c in range(1, max_col + 1) if c != no_col):
                ws.cell(r, no_col).value = None

    # ── Update Analysis sheet in place ────────────────────────────────────────
    try:
        analysis_name = _detect_sheet(wb, 'analysis')
    except KeyError:
        analysis_name = None

    pivot_year_col = fy_col or acc_col or sa_col
    net_source_col = net_col or amt_col

    if analysis_name and pivot_year_col and net_source_col and uc_col:
        ws_a = wb[analysis_name]
        # Safe mode: if we can't make sense of the Analysis layout, leave it
        # alone — the pivot's own refresh-on-open brings in the new rows,
        # and nothing of the accountant's own working is at risk.
        snapshot = _snapshot_sheet(ws_a)
        try:
            _update_analysis(ws_a, ws, pivot_year_col, net_source_col,
                             uc_col, last_data_row=new_last)
            problem = _verify_analysis_safe(ws_a, snapshot)
            if problem:
                _restore_sheet(ws_a, snapshot)
                print(f"  [excel_writer] Analysis layout unfamiliar ({problem}); "
                      f"leaving {analysis_name} untouched — pivot refresh on open "
                      f"will pick up the new rows.")
            else:
                print(f"  [excel_writer] Updated {analysis_name} in place.")

                # Grid's rightmost column: the farthest-right non-blank cell
                # in any of the pivot's own header rows.
                grid_last_col = 1
                for r in range(2, 6):
                    c = 2
                    while ws_a.cell(r, c).value not in (None, ''):
                        grid_last_col = max(grid_last_col, c)
                        c += 1
                _relocate_pivot_clear_of_grid(ws_a, grid_last_col)
        except Exception as exc:
            _restore_sheet(ws_a, snapshot)
            print(f"  [excel_writer] Analysis update failed ({exc}); {analysis_name} "
                  f"left untouched — pivot refresh on open will pick up new rows.")
    elif analysis_name:
        print("  [excel_writer] Missing year/net/category columns — Analysis left untouched.")

    _refresh_pivot_sources(wb, new_last, raw_sheet=ws.title)

    # New formula cells carry no saved result, so tell Excel to recalculate
    # the whole workbook when it opens — otherwise Excel for the web and
    # older versions can show them blank. (The template's own calculation
    # mode, e.g. manual, is left as it is.)
    if wb.calculation is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True

    wb.save(output_path)
    print(f"  [excel_writer] Saved -> {output_path}")

    problems = reconcile_balances(transactions, ws.cell(last_row, bal_col).value
                                  if bal_col and last_row > 1 else None)
    for p in problems:
        print(f"  [excel_writer] BALANCE WARNING: {p}")

    changed = _check_original_rows_untouched(template_path, output_path, ws.title,
                                             last_row)
    if changed:
        msg = f"Original rows modified unexpectedly: {changed[:3]}"
        print(f"  [excel_writer] INTEGRITY WARNING: {msg}")
        problems.append(msg)
    return problems


def _check_original_rows_untouched(template_path, output_path, sheet_name: str,
                                    last_row: int) -> list:
    """After writing, verify not a single original transaction cell changed.
    This is a hard guarantee for any template, however unusual: appending
    new rows must never touch what was already there."""
    orig = openpyxl.load_workbook(template_path)
    new = openpyxl.load_workbook(output_path)
    if sheet_name not in orig.sheetnames or sheet_name not in new.sheetnames:
        return []
    a, b = orig[sheet_name], new[sheet_name]
    diffs = []
    for r in range(1, last_row + 1):
        for c in range(1, a.max_column + 1):
            v1, v2 = a.cell(r, c).value, b.cell(r, c).value
            if v1 == v2:
                continue
            # openpyxl rounds floats on save (16.846310000000003 -> 16.84631);
            # that's a lossless re-print, not a data change.
            if (isinstance(v1, float) and isinstance(v2, float)
                    and abs(v1 - v2) < 1e-6):
                continue
            # row 1 is the header — a new 'Correct category' column is allowed.
            if r == 1 and v1 is None:
                continue
            diffs.append(f"{a.cell(r, c).coordinate}: {v1!r} -> {v2!r}")
    return diffs


def reconcile_balances(transactions: list, opening=None) -> list:
    """Check every balance against the one before it plus the row's amount,
    starting from the template's last balance. Returns readable problems
    (empty when everything adds up). A break means a missed, doubled or
    misread transaction somewhere — the accountant must be told."""
    problems = []
    prev = opening if isinstance(opening, (int, float)) and not isinstance(opening, bool) else None
    for t in transactions:
        bal = t.get('balance')
        if bal is None:
            prev = None                 # can't check across a blank balance
            continue
        if prev is not None and abs(round(prev + _signed(t), 2) - bal) >= 0.005:
            d = t['date'].date() if isinstance(t['date'], datetime) else t['date']
            problems.append(
                f"{d} {t.get('description') or ''} {_signed(t):+.2f}: balance {bal:.2f} "
                f"but previous {prev:.2f} {_signed(t):+.2f} = {prev + _signed(t):.2f}")
        prev = bal
    return problems


# ── Safe-mode guards: snapshot, restore, verify ──────────────────────────────

def _snapshot_sheet(ws) -> dict:
    """Capture every cell's value so we can roll back on an unsafe write."""
    return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row
            if c.value is not None}


def _restore_sheet(ws, snapshot: dict) -> None:
    """Put the sheet back exactly as _snapshot_sheet saw it. Cells the
    writer added that weren't in the snapshot are cleared."""
    current = {(c.row, c.column) for row in ws.iter_rows() for c in row
               if c.value is not None}
    for coord in current - snapshot.keys():
        ws.cell(*coord).value = None
    for (r, c), v in snapshot.items():
        ws.cell(r, c).value = v


def _verify_analysis_safe(ws, snapshot: dict) -> str | None:
    """Check that the Analysis update didn't destroy anything important.
    Returns a short reason to roll back, or None if the update looks fine.

    Red flags:
      - An original text value that wasn't a pivot header or category label
        got replaced with None or a different label (accountant's working lost).
      - A cell holding a formula was blanked (SUMIFs/GETPIVOTDATA destroyed).
      - A category label that was in the snapshot vanished entirely from
        column A (losing a BS/PL tag the accountant lined up next to it)."""
    labels_before = {str(v).strip() for (r, c), v in snapshot.items()
                     if c == 1 and isinstance(v, str) and v.strip()
                     and not v.startswith('=')
                     and str(v).strip().lower() not in {'row labels', 'grand total'}
                     and not str(v).lower().startswith('sum of')}
    labels_after = {str(ws.cell(r, 1).value or '').strip()
                    for r in range(1, ws.max_row + 1)
                    if ws.cell(r, 1).value not in (None, '')}
    lost = labels_before - labels_after
    if lost:
        return f'category label(s) vanished: {sorted(lost)[:3]}'

    # A formula cell blanked by the writer is almost always a bug.
    for (r, c), v in snapshot.items():
        if isinstance(v, str) and v.startswith('=') and ws.cell(r, c).value in (None, ''):
            # Allow the SUMIFS cells we deliberately rewrite (col 2+ in the grid)
            if c >= 2 and r >= 3:
                continue
            return f'formula at {ws.cell(r, c).coordinate} was lost'
    return None


# ── Analysis update (values only — formatting preserved) ─────────────────────

def _update_analysis(ws_a, raw_ws, year_col, net_col, uc_col, last_data_row) -> None:
    raw_name    = raw_ws.title
    year_letter = get_column_letter(year_col)
    net_letter  = get_column_letter(net_col)
    uc_letter   = get_column_letter(uc_col)

    data_cats = []              # RAW's own spelling, one per stripped name
    seen_cats = set()
    for r in range(2, last_data_row + 1):
        c = raw_ws.cell(r, uc_col).value
        if c is not None and str(c).strip() and not str(c).startswith('='):
            cs = str(c).strip()
            if cs not in seen_cats:
                seen_cats.add(cs)
                data_cats.append(str(c))

    # Year columns: the RAW year column's own labels. Row-4 cells only count
    # if they really are such labels — some templates keep the accountant's
    # side workings in that row (e.g. 50, 'PL', 'Grand Total', 45078), and
    # treating those as years gave a grid of SUMIFS that match nothing.
    # Existing year headers keep their order; years that appear in the data
    # but not yet on the sheet (e.g. FY26 after appending a new year) are
    # added as new columns, so the new year's totals actually show up.
    data_years = []
    for r in range(2, last_data_row + 1):
        v = raw_ws.cell(r, year_col).value
        if v not in (None, '') and not str(v).startswith('=') and v not in data_years:
            data_years.append(v)
    known = {str(y).strip() for y in data_years}

    # Only the unbroken run of year headers from column B is the grid.
    # Year labels further right are the accountant's own side workings (e.g.
    # 'BS'/'PL' tags in E, then their own 22/23 / 23/24 SUMIF summary in
    # G:H) — counting those made the grid swallow and wipe those columns.
    # A header also counts when it is shaped like a year label ('21/22',
    # 'FY23', 'SA24', 2023) even if no RAW row carries it any more —
    # otherwise one retired year at B4 would hide the whole grid.
    # Bare numbers like 50 (side workings) must NOT count, so a label needs
    # two year parts ('21/22'), a prefix ('FY23', 'SA24/25') or 19xx/20xx.
    year_like = re.compile(
        r'^(?:(?:FY|SA|YE)\s*\d{2}(?:\d{2})?(?:\s*[/.\-]\s*\d{2}(?:\d{2})?)?'
        r'|\d{2}(?:\d{2})?\s*[/\-]\s*\d{2}(?:\d{2})?'
        r'|(?:19|20)\d{2})$', re.I)
    # Templates vary in where the pivot header sits. Find the real header
    # row by looking for 'Row Labels' in column A within the first 10 rows:
    #   Ahmed:    row 3 = 'Sum of ...', row 4 = 'Row Labels' + years, cats 5+
    #   Chido:    row 2 = 'Sum of ...', row 3 = 'Row Labels' + years, cats 4+
    #   Kristaps: row 3 = 'Row Labels' + 'Sum of Net' (condensed),  cats 4+
    header_r = None
    caption_r = None
    for r in range(1, 11):
        v = str(ws_a.cell(r, 1).value or '').strip().lower()
        if v == 'row labels' and header_r is None:
            header_r = r
        if v.startswith('sum of') and caption_r is None:
            caption_r = r
    condensed = False
    if header_r is None:
        header_r = (caption_r or 3) + 1
    elif caption_r is None or caption_r == header_r:
        condensed = True
        caption_r = header_r
    if caption_r is None:
        caption_r = max(1, header_r - 1)
    first_cat_r = header_r + 1

    old_year_cols = []          # (col, value) of existing year headers
    for c in range(2, ws_a.max_column + 1):
        v = ws_a.cell(header_r, c).value
        if v is None or isinstance(v, float) or not (
                str(v).strip() in known or year_like.match(str(v).strip())):
            break
        old_year_cols.append((c, v))
    years = []
    for _, v in old_year_cols:
        if str(v).strip() not in {str(y).strip() for y in years}:
            years.append(v)
    have = {str(y).strip() for y in years}
    added = sorted((y for y in data_years if str(y).strip() not in have), key=str)

    # The old grid's extent — only this area is cleared, so anything else
    # the accountant keeps on the sheet (notes, side calculations) survives.
    # It spans the year headers plus any column still holding this tool's
    # own SUMIFS from an earlier run, down to the old 'Grand Total' row.
    grid_cols = max([c for c, _ in old_year_cols] + [1])
    sig = f"=SUMIFS('{raw_name}'!"
    for c in range(grid_cols + 1, ws_a.max_column + 1):
        if not str(ws_a.cell(first_cat_r, c).value or '').startswith(sig):
            break
        grid_cols = c
    if (str(ws_a.cell(header_r, grid_cols + 1).value or '').strip().lower() == 'grand total'):
        grid_cols += 1          # a pivot-style 'Grand Total' column

    grid_bottom = header_r
    for r in range(first_cat_r, ws_a.max_row + 1):
        v = ws_a.cell(r, 1).value
        if v in (None, ''):
            break
        grid_bottom = r
        if str(v).strip().lower() == 'grand total':
            break

    # Categories keep the rows they already have — accountants tag them in
    # the next column (e.g. 'PL' / 'BS' feeding their own SUMIFs), and
    # re-sorting would leave every tag against the wrong category. New
    # categories go at the bottom.
    # Labels are written back exactly as they were: SUMIFS matches text
    # exactly, so 'Professional fee' would miss RAW rows saying
    # 'Professional fee ' (trailing space) and drop them from the totals.
    old_cats = []
    for r in range(first_cat_r, grid_bottom + 1):
        v = ws_a.cell(r, 1).value
        if v not in (None, '') and str(v).strip().lower() != 'grand total':
            old_cats.append(v)
    old_keys = {str(c).strip() for c in old_cats}
    cats = (old_cats + sorted((c for c in data_cats if c.strip() not in old_keys), key=str.lower)
            if old_cats else sorted(data_cats, key=str.lower))

    # A new year (e.g. FY26) becomes a grid column only if those columns are
    # free. If the accountant uses them, the new year is left to the
    # PivotTable, which picks it up on refresh (see _refresh_pivot_sources).
    if added:
        first_new_col = 2 + len(years)
        need = range(first_new_col, first_new_col + len(added))
        busy = [f"{get_column_letter(c)}{r}" for c in need if c > grid_cols
                for r in range(3, first_cat_r + len(cats) + 1) if ws_a.cell(r, c).value not in (None, '')]
        if busy:
            print(f"  [excel_writer] Analysis: year(s) {added} not added as grid columns — "
                  f"cells {busy[:3]} are in use; they appear in the PivotTable instead.")
        else:
            years += added
            print(f"  [excel_writer] Analysis: added year column(s) {added}")

    # style models taken from the existing populated area
    hdr_style = copy(ws_a.cell(header_r, 2)._style)
    cat_style = copy(ws_a.cell(first_cat_r, 1)._style)
    val_style = copy(ws_a.cell(first_cat_r, 2)._style)

    caption = ws_a.cell(caption_r, 1).value or ('Row Labels' if condensed else 'Sum of NET')
    caption_b = ws_a.cell(caption_r, 2).value
    label_a = ws_a.cell(header_r, 1).value or 'Row Labels'

    # clear the old grid's values but keep every cell's formatting
    clear_from = caption_r
    for r in range(clear_from, grid_bottom + 1):
        for c in range(1, grid_cols + 1):
            ws_a.cell(r, c).value = None

    ws_a.cell(caption_r, 1, caption)
    if condensed:
        # single-row header: keep whatever the sum caption was (e.g. 'Sum of Net')
        if caption_b is not None:
            ws_a.cell(caption_r, 2, caption_b)
    else:
        if caption_b is not None:
            ws_a.cell(caption_r, 2, caption_b)
        else:
            ws_a.cell(caption_r, 2, 'Column Labels')
        ws_a.cell(header_r, 1, label_a)
        for j, yr in enumerate(years):
            cell = ws_a.cell(header_r, 2 + j)
            cell.value = yr
            cell._style = copy(hdr_style)

    for k, cat in enumerate(cats):
        r = first_cat_r + k
        cell = ws_a.cell(r, 1)
        cell.value = cat
        cell._style = copy(cat_style)
        if condensed and not years:
            # Single 'Sum of Net' column with no year axis — SUMIF total.
            # Essential: the pivot was moved out of A:B when it clashed with
            # the accountant's PL/BS summary, so these cells would freeze on
            # their old pivot values otherwise.
            formula = (
                f"=SUMIF('{raw_name}'!${uc_letter}:${uc_letter},$A{r},"
                f"'{raw_name}'!${net_letter}:${net_letter})"
            )
            cell = ws_a.cell(r, 2)
            cell.value = formula
            cell._style = copy(val_style)
        for j, yr in enumerate(years):
            yr_ref = f"${get_column_letter(2 + j)}${header_r}"
            formula = (
                f"=SUMIFS('{raw_name}'!${net_letter}:${net_letter},"
                f"'{raw_name}'!${year_letter}:${year_letter},{yr_ref},"
                f"'{raw_name}'!${uc_letter}:${uc_letter},$A{r})"
            )
            cell = ws_a.cell(r, 2 + j)
            cell.value = formula
            cell._style = copy(val_style)

    total_r = first_cat_r + len(cats)
    cell = ws_a.cell(total_r, 1)
    cell.value = 'Grand Total'
    cell._style = copy(cat_style)
    total_cols = len(years) if years else (1 if condensed else 0)
    for j in range(total_cols):
        col_letter = get_column_letter(2 + j)
        cell = ws_a.cell(total_r, 2 + j)
        cell.value = f"=SUM({col_letter}{first_cat_r}:{col_letter}{total_r - 1})"
        cell._style = copy(val_style)
