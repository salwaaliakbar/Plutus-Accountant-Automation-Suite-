"""Categorisation WITH the Claude API.

Two kinds of test:
 - mocked: a fake Anthropic client returns canned answers, so the batching,
   de-duplication, parsing and review flags are checked for free and always run;
 - live: a real call, run only when ANTHROPIC_API_KEY is set (marked `live`).
   Run just those with:  python -m pytest -m live
"""

import json
import os
from types import SimpleNamespace

import pytest

from core import categorise as cat

HAS_KEY = bool(os.environ.get('ANTHROPIC_API_KEY', '')) \
    and not os.environ['ANTHROPIC_API_KEY'].startswith('sk-ant-...')


def txn(desc, out=0.0, inn=0.0, **extra):
    return {'date': '2025-04-03', 'description': desc, 'money_in': inn,
            'money_out': out, **extra}


class FakeClient:
    """Stands in for anthropic.Anthropic. `answer(n_items, user_msg)` returns
    the text of Claude's reply."""

    def __init__(self, answer):
        self.calls = []
        self._answer = answer
        self.messages = SimpleNamespace(create=self._create)

    def _create(self, **kw):
        user_msg = kw['messages'][0]['content']
        self.calls.append({'system': kw['system'], 'user': user_msg})
        n = user_msg.count('\n') or 1
        return SimpleNamespace(content=[SimpleNamespace(type='text', text=self._answer(n, user_msg))])


@pytest.fixture
def fake_api(monkeypatch):
    def install(answer):
        client = FakeClient(answer)
        monkeypatch.setenv('ANTHROPIC_API_KEY', 'sk-ant-test-key')
        monkeypatch.setattr(cat.anthropic, 'Anthropic', lambda **kw: client)
        return client
    return install


def reply(*pairs):
    return lambda n, msg: json.dumps([{'category': c, 'sure': s} for c, s in pairs])


# ── mocked ───────────────────────────────────────────────────────────────────

def test_ai_answers_are_used_and_unsure_rows_flagged(fake_api):
    client = fake_api(reply(('Petrol', True), ('Software', False)))
    txns = [txn('SHELL PETROL STATION', out=45.5), txn('NEW SAAS TOOL', out=12)]
    cats, review = cat.categorise_detailed(txns)
    assert cats == ['Petrol', 'Software']
    assert review == [False, True]                   # not-sure answer is highlighted
    assert len(client.calls) == 1


def test_history_rows_never_reach_the_api(fake_api):
    client = fake_api(reply(('Petrol', True)))
    examples = [{'description': 'ACME OPTICAL SUPPLIES LTD', 'category': 'Purchase',
                 'direction': 'out'}] * 3
    txns = [txn('ACME OPTICAL SUPPLIES LTD', out=50), txn('SHELL PETROL STATION', out=45.5)]
    cats, _ = cat.categorise_detailed(txns, client_examples=examples)
    assert cats == ['Purchase', 'Petrol']
    assert 'ACME' not in client.calls[0]['user'].split('Categorise these transactions:')[1]


def test_identical_transactions_are_asked_once(fake_api):
    client = fake_api(reply(('Parking', True)))
    txns = [txn('DEMO HIGH ST PARKING', out=3.2)] * 3
    cats, _ = cat.categorise_detailed(txns)
    assert cats == ['Parking'] * 3
    assert client.calls[0]['user'].count('DEMO HIGH ST PARKING') == 1


def test_large_statements_are_sent_in_batches(fake_api):
    client = fake_api(lambda n, msg: json.dumps([{'category': 'Sundry', 'sure': True}] * 40))
    txns = [txn(f'PAYEE {chr(65 + i % 26)}{chr(65 + i // 26)}', out=1 + i) for i in range(50)]
    cats, _ = cat.categorise_detailed(txns, batch_size=40)
    assert len(client.calls) == 2 and len(cats) == 50


def test_client_spelling_wins_over_the_ais(fake_api):
    fake_api(reply(('DLA', True)))
    cats, _ = cat.categorise_detailed([txn('TRANSFER TO DIRECTOR LOAN', out=500)],
                                      analysis_vocab=['Directors Loan Account'])
    assert cats == ['Directors Loan Account']


def test_short_or_garbled_reply_degrades_to_unknown(fake_api):
    fake_api(reply(('Petrol', True)))                # 1 answer for 2 transactions
    cats, review = cat.categorise_detailed([txn('SHELL PETROL', out=1), txn('MYSTERY CO', out=2)])
    assert cats == ['Petrol', 'Unknown'] and review == [False, True]


def test_truncated_reply_keeps_the_complete_answers(fake_api):
    fake_api(lambda n, msg: '[{"category": "Petrol", "sure": true}, {"category": "Par')
    cats, _ = cat.categorise_detailed([txn('SHELL PETROL', out=1), txn('DEMO PARKING', out=2)])
    assert cats == ['Petrol', 'Unknown']


def test_sensitive_numbers_are_not_sent_to_the_api(fake_api):
    client = fake_api(reply(('Transfer', True)))
    cat.categorise_detailed([txn('To A/C 80522920 sort 12-34-56 IBAN GB29NWBK60161331926819',
                                 out=10, reference='ACC 12345678')])
    sent = client.calls[0]['user']
    for secret in ('80522920', '12-34-56', 'GB29NWBK60161331926819', '12345678'):
        assert secret not in sent


def test_client_vocabulary_and_examples_are_in_the_prompt(fake_api):
    client = fake_api(reply(('Income', True)))
    examples = [{'description': 'EXAMPLE NHS PAYMENTS', 'category': 'PCSE income', 'direction': 'in'}]
    cat.categorise_detailed([txn('SOME NEW PAYER', inn=100)], client_examples=examples)
    assert "'PCSE income'" in client.calls[0]['system']


def test_no_key_marks_everything_unknown_for_review(monkeypatch):
    monkeypatch.setenv('ANTHROPIC_API_KEY', '')
    cats, review = cat.categorise_detailed([txn('MYSTERY CO', out=2)])
    assert cats == ['Unknown'] and review == [True]


# ── live (real Claude call) ──────────────────────────────────────────────────

@pytest.mark.live
@pytest.mark.skipif(not HAS_KEY, reason='ANTHROPIC_API_KEY not set')
def test_live_claude_categorises_obvious_payees():
    txns = [
        txn('SHELL PETROL STATION', out=45.50, subcategory='Card payment'),
        txn('ASSOCIATION OF OPTOMETRISTS', out=250, subcategory='Direct debit'),
        txn('SPECSAVERS OPTICIANS', inn=3200, subcategory='BACS credit'),
        txn('HMRC VAT', out=1800, reference='VAT 03 25', subcategory='Faster payment'),
    ]
    cats, review = cat.categorise_detailed(txns)
    assert len(cats) == 4 and 'Unknown' not in cats
    assert cats[0] == 'Petrol'
    assert cats[1] in ('AOP', 'Professional - AOP', 'Insurance - AOP')
    assert cats[2] == 'Income'
    assert cats[3].startswith('HMRC')


def test_reply_wrapped_in_code_fence_keeps_every_answer(fake_api):
    # Claude often wraps the array in ```json ... ``` — the last answer used to be lost
    fake_api(lambda n, msg: '```json\n[{"category": "Petrol", "sure": true}, '
                            '{"category": "HMRC-VAT", "sure": true}]\n```\nHope that helps!')
    cats, review = cat.categorise_detailed([txn('SHELL PETROL', out=1), txn('HMRC VAT', out=2)])
    assert cats == ['Petrol', 'HMRC-VAT'] and review == [False, False]
