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
