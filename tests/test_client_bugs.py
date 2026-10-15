"""Regression tests for the bugs found in client testing (round 2).
All data is invented."""

import time
from datetime import date, datetime

import openpyxl
import pytest

from core import categorise as cat
from parsers.parse_csv import parse_csv
from parsers.parse_xls import parse_xls

# ── XLSX dates saved with day and month swapped ──────────────────────────────

ROWS = [  # (true date, description, amount, balance) — oldest first, like the bank's file
    (date(2025, 3, 2), 'SUBWAY', -9.0, 991.0), (date(2025, 3, 3), 'SPECSAVERS', 500.0, 1491.0),
    (date(2025, 3, 5), 'BOOTS', -11.0, 1480.0), (date(2025, 3, 7), 'PARKING', -3.0, 1477.0),
    (date(2025, 3, 11), 'COFFEE', -4.0, 1473.0), (date(2025, 3, 13), 'SPECSAVERS', 300.0, 1773.0),
    (date(2025, 4, 1), 'COFFEE', -4.5, 1768.5), (date(2025, 4, 8), 'SHELL', -40.0, 1728.5),
    (date(2025, 4, 22), 'TESCO', -12.0, 1716.5), (date(2025, 5, 4), 'SPECSAVERS', 250.0, 1966.5),
]


def swapped_xlsx(tmp_path):
    """What a US-locale Excel makes of a day-first CSV: days <= 12 become real
    dates with day and month exchanged, days > 12 stay as correct text."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Date', 'Counter Party', 'Amount (GBP)', 'Balance (GBP)'])
    for d, desc, amt, bal in ROWS:
        if d.day <= 12 and d.day != d.month:
            shown = datetime(d.year, d.day, d.month)            # swapped
        elif d.day <= 12:
            shown = datetime(d.year, d.month, d.day)            # day == month: identical either way
        else:
            shown = d.strftime('%d/%m/%Y')                      # text, correct
        ws.append([shown, desc, amt, bal])
    p = tmp_path / 'swapped.xlsx'
    wb.save(p)
    return p


def test_swapped_dates_are_put_back(tmp_path):
    txns = parse_xls(swapped_xlsx(tmp_path))
    assert [t['date'] for t in txns] == [r[0] for r in ROWS]
    assert [t['balance'] for t in txns] == [r[3] for r in ROWS]


def test_correct_dates_are_left_alone(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Date', 'Description', 'Amount', 'Balance'])
    for d, desc, amt, bal in ROWS:
        ws.append([datetime(d.year, d.month, d.day), desc, amt, bal])
    p = tmp_path / 'fine.xlsx'
    wb.save(p)
    assert [t['date'] for t in parse_xls(p)] == [r[0] for r in ROWS]


def test_newest_first_file_with_correct_dates_is_not_mistaken_for_swapped(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Date', 'Description', 'Amount', 'Balance'])
    for d, desc, amt, bal in reversed(ROWS):
        ws.append([datetime(d.year, d.month, d.day), desc, amt, bal])
    p = tmp_path / 'newest_first.xlsx'
    wb.save(p)
    assert [t['date'] for t in parse_xls(p)] == [r[0] for r in ROWS]


def test_mojibake_in_a_spreadsheet_is_repaired(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Date', 'Description', 'Amount'])
    ws.append([datetime(2025, 5, 1), 'CaffÃ¨ Nero', -3.2])
    p = tmp_path / 'm.xlsx'
    wb.save(p)
    assert parse_xls(p)[0]['description'] == 'Caffè Nero'


# ── PDF with Type / Transaction / Category / Value columns (Starling report) ─

reportlab = pytest.importorskip('reportlab')


def test_pdf_with_value_and_category_columns(tmp_path, monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')      # never fall through to the real AI parser
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from parsers.parse_pdf import parse_pdf

    rows = [('01/12/2025', 'Direct Credit', 'Example Optics Pr (1003', 'Revenue', '+£2,860.00', '+£3,860.00'),
            ('01/12/2025', 'Direct Debit', 'Demo Finance', 'Vehicles', '-£464.98', '+£3,395.02'),
            ('05/12/2025', 'Faster Payment', 'Sample Person', 'None', '-£500.00', '+£2,895.02')]
    p = tmp_path / 'report.pdf'
    c = canvas.Canvas(str(p), pagesize=A4)
    c.setFont('Helvetica', 7)
    y = 780
    c.drawString(40, y, 'Transaction Export 01 Dec 2025 to 05 Jan 2026')
    y -= 30
    for x, h in ((38, 'DATE'), (101, 'TYPE'), (191, 'TRANSACTION'), (321, 'CATEGORY'),
                 (438, 'VALUE'), (513, 'BALANCE')):
        c.drawString(x, y, h)
    for d, typ, desc, catg, val, bal in rows:
        y -= 15            # the row sits right under the header, like the real file
        c.drawString(38, y, d); c.drawString(101, y, typ); c.drawString(191, y, desc)
        c.drawString(321, y, catg); c.drawString(417, y, val); c.drawString(500, y, bal)
    c.save()

    txns = parse_pdf(p, year_hint=2025)
    assert [(t['money_in'], t['money_out']) for t in txns] == [(2860.0, 0), (0, 464.98), (0, 500.0)]
    assert [t['subcategory'] for t in txns] == ['Direct Credit', 'Direct Debit', 'Faster Payment']
    assert [t['csv_category'] for t in txns] == ['Revenue', 'Vehicles', '']    # 'None' means no category
    assert txns[1]['description'] == 'Demo Finance'


# ── Categorisation ────────────────────────────────────────────────────────────

def test_card_payments_in_different_months_are_the_same_payee():
    a = cat._payee_key('4445 09MAY23 CD , MCDONALDS 1194 , FROME GB')
    b = cat._payee_key('4271 16FEB26 , MCDONALDS 1194 , FROME GB')
    assert a == b and a.startswith('MCDONALDS')


def test_history_now_matches_a_card_payee_in_a_new_month(monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    history = [{'description': f'4445 {d}JUN23 CD , UBER *EATS , HELP.UBER.COM GB',
                'category': 'Lunch', 'direction': 'out'} for d in ('09', '12', '18')]
    t = {'description': '8057 17SEP25 , UBER *EATS , HELP.UBER.COM GB', 'money_in': 0, 'money_out': 20}
    cats, review = cat.categorise_detailed([t], client_examples=history)
    assert cats == ['Lunch'] and review == [False]


@pytest.mark.parametrize('desc, ref, expected', [
    ('HMRC Cumbernauld', '120PP034375842609', 'HMRC-PAYE'),
    ('HMRC Cumbernauld', '4608109978A00101A', 'HMRC-CT'),
    ('HMRC Cumbernauld', '7964830073K', 'HMRC-SA'),
    ('Locums Limited', '120PN03480909', None),         # same shape, but not HMRC
    ('HMRC NDDS', '', None),                           # no reference: the AI / accountant decides
])
def test_hmrc_payments_are_recognised_by_reference(desc, ref, expected):
    assert cat._hmrc_by_reference({'description': desc, 'reference': ref, 'money_out': 10}) == expected


def test_hmrc_refund_is_not_forced_into_a_tax(monkeypatch):
    assert cat._hmrc_by_reference({'description': 'HMRC', 'reference': '7964830073K',
                                   'money_in': 10, 'money_out': 0}) is None


def test_hmrc_reference_uses_the_clients_spelling(monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    t = {'description': 'HMRC Cumbernauld', 'reference': '120PP034375842609', 'money_in': 0, 'money_out': 900}
    cats, review = cat.categorise_detailed([t], analysis_vocab=['HMRC - PAYE'])
    assert cats == ['HMRC - PAYE'] and review == [False]


# ── API: every category is reported, with template warnings ──────────────────

def test_api_reports_every_category_and_template_warnings(monkeypatch, tmp_path, make_csv, make_template):
    pytest.importorskip('httpx')
    from fastapi.testclient import TestClient
    from api import server

    out, up = tmp_path / 'output', tmp_path / 'uploads'
    out.mkdir(); up.mkdir()
    monkeypatch.setattr(server, 'JOBS', server.JobStore(out / 'jobs', out, up))
    monkeypatch.setattr(server, 'OUTPUT_DIR', out)
    monkeypatch.setattr(server, 'UPLOAD_DIR', up)
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    monkeypatch.setattr(server, '_run_pipeline', server._run_pipeline)

    # six different categories from the (mocked) categoriser
    import core.categorise as c
    monkeypatch.setattr(c, 'categorise_detailed',
                        lambda txns, **kw: (['A', 'B', 'C', 'D', 'E'][:len(txns)], [False] * len(txns)))

    client = TestClient(server.app)
    path, tmpl = make_csv(), make_template()
    r = client.post('/api/upload', files={'file': (path.name, path.read_bytes()),
                                          'template': ('t.xlsx', tmpl.read_bytes())})
    job_id = r.json()['job_id']
    end = time.time() + 20
    while time.time() < end:
        job = client.get(f'/api/progress/{job_id}').json()
        if job['status'] != 'processing':
            break
        time.sleep(0.1)
    assert job['status'] == 'done', job['message']
    assert job['category_count'] == 5 == len(job['top_categories'])      # was capped at 5 total
    assert isinstance(job['template_warnings'], list)


def test_api_says_clearly_when_everything_is_already_in_the_workbook(monkeypatch, tmp_path, make_csv, make_template):
    pytest.importorskip('httpx')
    from fastapi.testclient import TestClient
    from api import server
    from core.excel_writer import prepare_new_transactions, write_workbook

    out, up = tmp_path / 'output', tmp_path / 'uploads'
    out.mkdir(); up.mkdir()
    monkeypatch.setattr(server, 'JOBS', server.JobStore(out / 'jobs', out, up))
    monkeypatch.setattr(server, 'OUTPUT_DIR', out)
    monkeypatch.setattr(server, 'UPLOAD_DIR', up)

    # template that already contains the statement
    txns, _ = prepare_new_transactions(parse_csv(make_csv()), make_template())
    full = tmp_path / 'full.xlsx'
    write_workbook(txns, ['X'] * len(txns), make_template(), full)

    client = TestClient(server.app)
    path = make_csv()
    r = client.post('/api/upload', files={'file': (path.name, path.read_bytes()),
                                          'template': ('t.xlsx', full.read_bytes())})
    job_id = r.json()['job_id']
    end = time.time() + 20
    while time.time() < end:
        job = client.get(f'/api/progress/{job_id}').json()
        if job['status'] != 'processing':
            break
        time.sleep(0.1)
    assert job['status'] == 'error'
    assert 'already in the workbook' in job['message'] and '5 transactions' in job['message']


# ── Monzo-style statements: the merchant is in 'Name', not 'Description' ─────

def test_payee_column_is_kept_separately(tmp_path):
    csv = tmp_path / 'monzo.csv'
    csv.write_text('Transaction ID,Date,Type,Name,Description,Amount,Balance\n'
                   'tx1,03/04/2025,Direct Debit,EE,Q53133604629274491,-25.00,975.00\n',
                   encoding='utf-8')
    (t,) = parse_csv(csv)
    assert t['description'] == 'Q53133604629274491' and t['payee'] == 'EE'


def test_history_matches_on_the_payee_when_the_description_is_a_code(monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    history = [{'description': f'Q5313360462927449{i}', 'payee': 'EE', 'category': 'Mobile',
                'direction': 'out'} for i in range(3)]
    t = {'description': 'Q53133604714935709', 'payee': 'EE', 'money_in': 0, 'money_out': 25}
    cats, review = cat.categorise_detailed([t], client_examples=history)
    assert cats == ['Mobile'] and review == [False]


def test_payee_is_sent_to_the_ai_when_the_description_does_not_name_it():
    assert cat._payee_text({'description': 'Q531336', 'payee': 'EE'}) == 'EE Q531336'
    assert cat._payee_text({'description': 'TESCO STORES 2527', 'payee': 'Tesco'}) == 'TESCO STORES 2527'


def test_template_name_column_is_found_and_filled(tmp_path, make_template):
    from core.columns import resolve_columns
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'RAW'
    ws.append(['No', 'Date', 'Name', 'Description', 'Money In', 'Money Out', 'Balance', 'UC category'])
    hdr = {ws.cell(1, c).value: c for c in range(1, 9)}
    rc = resolve_columns(ws)
    assert rc['desc'] == hdr['Description'] and rc['name'] == hdr['Name']


# ── 'Food?' is the accountant's query, not a category ────────────────────────

def test_question_mark_categories_become_the_category_plus_a_review_flag(monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    history = [{'description': 'GREGGS PLC', 'category': 'Food?', 'direction': 'out'}] * 3
    t = {'description': 'GREGGS PLC', 'money_in': 0, 'money_out': 3}
    cats, review = cat.categorise_detailed([t], client_examples=history)
    assert cats == ['Food'] and review == [True]


def test_question_mark_spellings_are_not_offered_to_the_ai():
    lines, client_cats = cat._build_examples(
        [{'description': 'GREGGS PLC', 'category': 'Food?'},
         {'description': 'SHELL', 'category': 'Petrol'}])
    assert client_cats == ['Food', 'Petrol']
    assert all('?' not in line for line in lines)
