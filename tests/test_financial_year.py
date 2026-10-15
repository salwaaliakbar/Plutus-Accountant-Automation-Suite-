from datetime import date

import pytest

from core.financial_year import get_fy


@pytest.mark.parametrize('d, expected', [
    (date(2024, 5, 2),   ('FY24', '24/25')),
    (date(2024, 4, 5),   ('FY23', '23/24')),   # last day of the old year
    (date(2024, 4, 6),   ('FY24', '24/25')),   # first day of the new year
    (date(2025, 1, 1),   ('FY24', '24/25')),
    (date(1999, 12, 31), ('FY99', '99/00')),
])
def test_uk_tax_year_boundary(d, expected):
    assert get_fy(d) == expected
