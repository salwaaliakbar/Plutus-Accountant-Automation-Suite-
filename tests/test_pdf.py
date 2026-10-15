import pytest

pytest.importorskip('reportlab')
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from conftest import STATEMENT
from parsers.parse_pdf import parse_pdf


def _text_pdf(path):
    c = canvas.Canvas(str(path), pagesize=A4)
    y = 800
    c.drawString(40, y, 'Demo Bank - Statement of account')
    y -= 30
    for x, h in ((40, 'Date'), (130, 'Description'), (360, 'Paid in'),
                 (430, 'Paid out'), (500, 'Balance')):
        c.drawString(x, y, h)
    for d, desc, _t, pin, pout, bal in STATEMENT:
        y -= 18
        c.drawString(40, y, d.strftime('%d/%m/%Y'))
        c.drawString(130, y, desc)
        if pin:
            c.drawRightString(405, y, f'{pin:.2f}')
        if pout:
            c.drawRightString(480, y, f'{pout:.2f}')
        c.drawRightString(550, y, f'{bal:.2f}')
    c.save()


def test_text_pdf_without_grid_lines(tmp_path):
    p = tmp_path / 'statement.pdf'
    _text_pdf(p)
    txns = parse_pdf(p, year_hint=2025)
    assert len(txns) == len(STATEMENT)
    assert [(t['money_in'], t['money_out']) for t in txns] == \
        [(pin, pout) for *_, pin, pout, _b in STATEMENT]


def test_image_only_pdf_is_reported_as_scanned(tmp_path, monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')       # no key: the scan cannot be read
    # a page with a drawn shape but no text layer, like a scan
    p = tmp_path / 'scan.pdf'
    c = canvas.Canvas(str(p), pagesize=A4)
    c.rect(50, 50, 400, 600, fill=1)
    c.save()
    with pytest.raises(ValueError, match='scanned image'):
        parse_pdf(p, year_hint=2025)
