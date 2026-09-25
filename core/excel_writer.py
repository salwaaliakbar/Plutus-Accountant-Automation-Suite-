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
from openpyxl.utils import column_index_from_string, get_column_letter

from core.financial_year import get_fy
from parsers._common import norm_header
from core.year_labels import YearLabeller, learn_year_columns


# ── Sheet / header helpers ────────────────────────────────────────────────────

def _detect_sheet(wb: openpyxl.Workbook, keyword: str) -> str:
    kw = keyword.lower()
    matches = [name for name in wb.sheetnames if kw in name.lower()]
    if not matches:
        raise KeyError(f"No sheet containing '{keyword}' found in {wb.sheetnames}")
    if kw == 'raw':
        # a sheet like 'Analysis 24 Raw (2)' is an analysis sheet, not the RAW tab
        pure = [m for m in matches if 'analysis' not in m.lower()]
        if pure:
            return pure[0]
    return matches[0]


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


def _refresh_pivot_sources(wb, new_last_row: int) -> None:
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
            m = _REF_RE.match(wsrc.ref or '')
            if not m:
                continue
            col1, row1, col2, row2 = m.groups()
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
    uc_col   = _col(hdrs, 'UC category', 'UC Category')
    ref_col  = _col(hdrs, 'Reference', 'Ref')
    type_col = _col(hdrs, 'Type', 'Transaction Type', 'Transaction type')

    if not desc_col or not uc_col:
        return []

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
        })
    return examples


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
) -> None:
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
    uc_col    = _col(hdrs, 'UC category', 'UC Category')
    ts_col     = _col(hdrs, 'Timestamp')
    from_col   = _col(hdrs, 'From')
    to_col     = _col(hdrs, 'To')
    status_col = _col(hdrs, 'Status')
    tag_col    = _col(hdrs, 'Tag 1', 'Tag')

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

    # ── Write transaction rows ────────────────────────────────────────────────
    for i, (txn, cat) in enumerate(zip(transactions, categories)):
        r = last_row + 1 + i
        seq = last_no + 1 + i

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
        w(csv_cat_col, (txn.get('csv_category') or '').strip() or None)
        w(uc_col,   cat)
        w(ts_col,     txn.get('timestamp'))
        w(from_col,   (txn.get('from')   or '').strip() or None)
        w(to_col,     (txn.get('to')     or '').strip() or None)
        w(status_col, (txn.get('status') or '').strip() or None)
        w(tag_col,    (txn.get('tag')    or '').strip() or None)

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
        _update_analysis(ws_a, ws, pivot_year_col, net_source_col,
                         uc_col, last_data_row=new_last)
        print(f"  [excel_writer] Updated {analysis_name} in place.")

        grid_last_col = 1
        c = 2
        while ws_a.cell(4, c).value not in (None, ''):
            grid_last_col = c
            c += 1
        _relocate_pivot_clear_of_grid(ws_a, grid_last_col)
    elif analysis_name:
        print("  [excel_writer] Missing year/net/category columns — Analysis left untouched.")

    _refresh_pivot_sources(wb, new_last)

    wb.save(output_path)
    print(f"  [excel_writer] Saved -> {output_path}")


# ── Analysis update (values only — formatting preserved) ─────────────────────

def _update_analysis(ws_a, raw_ws, year_col, net_col, uc_col, last_data_row) -> None:
    raw_name    = raw_ws.title
    year_letter = get_column_letter(year_col)
    net_letter  = get_column_letter(net_col)
    uc_letter   = get_column_letter(uc_col)

    data_cats = []
    for r in range(2, last_data_row + 1):
        c = raw_ws.cell(r, uc_col).value
        if c is not None and str(c).strip() and not str(c).startswith('='):
            cs = str(c).strip()
            if cs not in data_cats:
                data_cats.append(cs)

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

    old_year_cols = []          # (col, value) of existing year headers
    for c in range(2, ws_a.max_column + 1):
        v = ws_a.cell(4, c).value
        if v is not None and str(v).strip() in known:
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
    for c in range(2, ws_a.max_column + 1):
        if str(ws_a.cell(5, c).value or '').startswith(sig):
            grid_cols = max(grid_cols, c)
    if (str(ws_a.cell(4, grid_cols + 1).value or '').strip().lower() == 'grand total'):
        grid_cols += 1          # a pivot-style 'Grand Total' column
    grid_bottom = 4
    for r in range(5, ws_a.max_row + 1):
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
    old_cats = []
    for r in range(5, grid_bottom + 1):
        v = ws_a.cell(r, 1).value
        if v not in (None, '') and str(v).strip().lower() != 'grand total':
            old_cats.append(str(v).strip())
    cats = old_cats + sorted((c for c in data_cats if c not in old_cats), key=str.lower)         if old_cats else sorted(data_cats, key=str.lower)

    # A new year (e.g. FY26) becomes a grid column only if those columns are
    # free. If the accountant uses them, the new year is left to the
    # PivotTable, which picks it up on refresh (see _refresh_pivot_sources).
    if added:
        first_new_col = 2 + len(years)
        need = range(first_new_col, first_new_col + len(added))
        busy = [f"{get_column_letter(c)}{r}" for c in need if c > grid_cols
                for r in range(3, 6 + len(cats)) if ws_a.cell(r, c).value not in (None, '')]
        if busy:
            print(f"  [excel_writer] Analysis: year(s) {added} not added as grid columns — "
                  f"cells {busy[:3]} are in use; they appear in the PivotTable instead.")
        else:
            years += added
            print(f"  [excel_writer] Analysis: added year column(s) {added}")

    # style models taken from the existing populated area
    hdr_style = copy(ws_a.cell(4, 2)._style)
    cat_style = copy(ws_a.cell(5, 1)._style)
    val_style = copy(ws_a.cell(5, 2)._style)

    # clear the old grid's values but keep every cell's formatting
    for r in range(3, grid_bottom + 1):
        for c in range(1, grid_cols + 1):
            ws_a.cell(r, c).value = None

    ws_a.cell(3, 1, 'Sum of NET')
    ws_a.cell(3, 2, 'Column Labels')
    ws_a.cell(4, 1, 'Row Labels')
    for j, yr in enumerate(years):
        cell = ws_a.cell(4, 2 + j)
        cell.value = yr
        cell._style = copy(hdr_style)

    for k, cat in enumerate(cats):
        r = 5 + k
        cell = ws_a.cell(r, 1)
        cell.value = cat
        cell._style = copy(cat_style)
        for j, yr in enumerate(years):
            yr_ref = f"${get_column_letter(2 + j)}$4"
            formula = (
                f"=SUMIFS('{raw_name}'!${net_letter}:${net_letter},"
                f"'{raw_name}'!${year_letter}:${year_letter},{yr_ref},"
                f"'{raw_name}'!${uc_letter}:${uc_letter},$A{r})"
            )
            cell = ws_a.cell(r, 2 + j)
            cell.value = formula
            cell._style = copy(val_style)

    total_r = 5 + len(cats)
    cell = ws_a.cell(total_r, 1)
    cell.value = 'Grand Total'
    cell._style = copy(cat_style)
    for j in range(len(years)):
        col_letter = get_column_letter(2 + j)
        cell = ws_a.cell(total_r, 2 + j)
        cell.value = f"=SUM({col_letter}5:{col_letter}{total_r - 1})"
        cell._style = copy(val_style)
