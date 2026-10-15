"""Scanned (image-only) PDFs are read by Claude from the page images."""

import json
import os
from types import SimpleNamespace

import pytest
from dotenv import load_dotenv

load_dotenv()
pytest.importorskip('reportlab')
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from conftest import STATEMENT
from parsers import parse_pdf as P

HAS_KEY = bool(os.environ.get('ANTHROPIC_API_KEY', '')) \
    and not os.environ['ANTHROPIC_API_KEY'].startswith('sk-ant-...')


def text_pdf(path):
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


@pytest.fixture
def scanned_pdf(tmp_path):
    """A text statement photographed: rendered to an image and saved as an
    image-only PDF, so there is no text layer at all."""
    import pdfplumber
    src = tmp_path / 'source.pdf'
    text_pdf(src)
    with pdfplumber.open(src) as pdf:
        img = pdf.pages[0].to_image(resolution=130).original.convert('RGB')
    out = tmp_path / 'scan.pdf'
    img.save(out, 'PDF')
    return out


def test_fixture_really_has_no_text_layer(scanned_pdf):
    import pdfplumber
    with pdfplumber.open(scanned_pdf) as pdf:
        assert P._is_scanned(pdf)


class FakeClient:
    def __init__(self, reply):
        self.calls = []
        self.messages = SimpleNamespace(create=self._create)
        self._reply = reply

    def _create(self, **kw):
        self.calls.append(kw)
        return SimpleNamespace(content=[SimpleNamespace(type='text', text=self._reply)])


def statement_json():
    return json.dumps([
        {'date': d.isoformat(), 'description': desc, 'type': typ,
         'money_in': pin, 'money_out': pout, 'balance': bal}
        for d, desc, typ, pin, pout, bal in STATEMENT])


def test_scan_is_sent_as_page_images_and_parsed(scanned_pdf, monkeypatch):
    client = FakeClient('```json\n' + statement_json() + '\n```')
    monkeypatch.setenv('ANTHROPIC_API_KEY', 'sk-ant-test')
    monkeypatch.setattr(P, '_parse_scanned', P._parse_scanned)
    import anthropic
    monkeypatch.setattr(anthropic, 'Anthropic', lambda **kw: client)

    txns = P.parse_pdf(scanned_pdf, year_hint=2025)

    assert len(txns) == len(STATEMENT)
    assert all(t['from_scan'] for t in txns)
    assert [(t['money_in'], t['money_out']) for t in txns] == \
        [(pin, pout) for *_, pin, pout, _b in STATEMENT]
    blocks = client.calls[0]['messages'][0]['content']
    assert blocks[0]['type'] == 'image' and blocks[0]['source']['media_type'] == 'image/png'


def test_scan_reading_can_be_switched_off(scanned_pdf, monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', 'sk-ant-test')
    monkeypatch.setenv('PLUTUS_SCAN_AI', '0')
    with pytest.raises(ValueError, match='switched off'):
        P.parse_pdf(scanned_pdf, year_hint=2025)


def test_scan_without_api_key_says_why(scanned_pdf, monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    with pytest.raises(ValueError, match='API key'):
        P.parse_pdf(scanned_pdf, year_hint=2025)


def test_unreadable_scan_is_an_error_not_an_empty_result(scanned_pdf, monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', 'sk-ant-test')
    import anthropic
    monkeypatch.setattr(anthropic, 'Anthropic', lambda **kw: FakeClient('[]'))
    with pytest.raises(ValueError, match='No transactions could be read'):
        P.parse_pdf(scanned_pdf, year_hint=2025)


@pytest.mark.live
@pytest.mark.skipif(not HAS_KEY, reason='ANTHROPIC_API_KEY not set')
def test_live_claude_reads_a_scanned_statement(scanned_pdf, monkeypatch):
    monkeypatch.delenv('PLUTUS_SCAN_AI', raising=False)
    txns = P.parse_pdf(scanned_pdf, year_hint=2025)
    got = [(t['date'], t['money_in'], t['money_out'], t['balance']) for t in txns]
    want = [(d, pin, pout, bal) for d, _desc, _t, pin, pout, bal in STATEMENT]
    assert got == want
