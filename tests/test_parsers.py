import pytest

from conftest import STATEMENT
from parsers.parse_csv import parse_csv
from parsers.parse_xls import parse_xls


def simplify(txns):
    return [(t['date'].date() if hasattr(t['date'], 'date') else t['date'],
             t['description'], round(t['money_in'], 2), round(t['money_out'], 2))
            for t in txns]


EXPECTED = [(d, desc, pin, pout) for d, desc, _t, pin, pout, _b in STATEMENT]


def test_csv_paid_in_paid_out(make_csv):
    assert simplify(parse_csv(make_csv())) == EXPECTED


def test_csv_signed_amount_column(make_csv):
    assert simplify(parse_csv(make_csv(signed=True))) == EXPECTED


def test_csv_header_block_above_headings(make_csv):
    assert simplify(parse_csv(make_csv(header_block=True))) == EXPECTED


def test_csv_newest_first_comes_out_oldest_first_in_original_order(make_csv):
    # same-day rows must keep their order, so the balance check still works
    txns = parse_csv(make_csv(newest_first=True))
    assert simplify(txns) == EXPECTED
    assert [t['balance'] for t in txns] == [b for *_, b in STATEMENT]


def test_csv_semicolon_delimiter(tmp_path):
    p = tmp_path / 's.csv'
    p.write_text('Date;Description;Paid in;Paid out\n03/04/2025;ACME LTD;;10.50\n',
                 encoding='utf-8')
    (t,) = parse_csv(p)
    assert (t['money_in'], t['money_out']) == (0, 10.5)


def test_xlsx_matches_csv(make_csv, make_xlsx):
    assert simplify(parse_xls(make_xlsx())) == simplify(parse_csv(make_csv()))


def test_xlsx_cover_sheet_and_header_block(make_xlsx):
    assert simplify(parse_xls(make_xlsx(header_block=True, cover_sheet=True))) == EXPECTED


def test_legacy_xls_matches_csv(make_csv, make_xls):
    assert simplify(parse_xls(make_xls())) == simplify(parse_csv(make_csv()))


def test_repeated_identical_payments_are_kept(make_csv):
    descs = [t['description'] for t in parse_csv(make_csv())]
    assert descs.count('SHELL PETROL STATION') == 2


# ── A new bank's column wording colliding with the header-matching rules ─────

def test_amount_in_gbp_is_recognised_as_the_amount_column(tmp_path):
    # Mettle: 'Amount in GBP' must not be left as 'amount in' after the
    # currency suffix is stripped, or it fails to match 'Amount'.
    p = tmp_path / 'mettle.csv'
    p.write_text('Date,Amount in GBP,Balance,Reference,Description,Transaction Type,'
                 'Invoices,Receipts,Note,Category\n'
                 '03/04/2025,-20.35,980.00,Ringgo,Ringgo,Card purchase,,,,Motor expenses\n'
                 '04/04/2025,424.00,1404.00,Asda,Asda,Scheduled payment,,,,Sales\n',
                 encoding='utf-8')
    txns = parse_csv(p)
    assert [(t['money_in'], t['money_out']) for t in txns] == [(0, 20.35), (424.0, 0)]


def test_receipts_column_is_not_mistaken_for_money_in(tmp_path):
    # Mettle's 'Receipts' holds uploaded receipt filenames, not amounts — it
    # must not shadow a real Amount/Paid-in column by matching first.
    p = tmp_path / 'mettle2.csv'
    p.write_text('Date,Amount in GBP,Reference,Description,Receipts\n'
                 '03/04/2025,-20.35,Ringgo,Ringgo,\n'
                 '04/04/2025,-9.04,Tim Hortons,Tim Hortons,receipt-photo.jpeg\n',
                 encoding='utf-8')
    txns = parse_csv(p)
    assert [t['money_out'] for t in txns] == [20.35, 9.04]


def test_wrongly_matched_money_column_raises_instead_of_silently_empty(tmp_path):
    # If every dated row comes back with zero money, that is almost always
    # the wrong column being matched, not a genuinely empty statement — this
    # must be a clear, actionable error, not a silent empty result.
    p = tmp_path / 'bad.csv'
    p.write_text('Date,Description,Paid in,Note\n'
                 '03/04/2025,ACME LTD,,some note\n'
                 '04/04/2025,OTHER LTD,,another note\n',
                 encoding='utf-8')
    with pytest.raises(ValueError, match="none had a usable amount"):
        parse_csv(p)
