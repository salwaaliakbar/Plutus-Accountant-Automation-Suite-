"""Learn year-label conventions from a client's existing template rows.

Every client uses a different fiscal-year convention:
  - Farris:  FY25 = Apr 2024 – Mar 2025 (end-year label, April year)
  - Chido:   FY24 = Dec 2023 – Nov 2024 (company year ends 30 November!)
  - Ahmed:   '2025' = Apr 2024 – Mar 2025 (text label, end-year)

Instead of hardcoding any calendar rule, we learn the mapping from the
client's own data: group existing rows by label, anchor on the group with
the latest dates, and extrapolate labels forward by anniversaries of the
anchor group's first date.
"""

import re
from datetime import date, datetime, timedelta

# A fiscal year is at most ~12 months.  Rows further than this beyond a
# group's first date are mislabelled stragglers and must not stretch the
# group's range (client templates really do contain such rows).
_MAX_GROUP_SPAN = timedelta(days=365)

# Groups whose start day differs by no more than this (circular days) are
# treated as observations of the same fiscal-year start day.
_BOUNDARY_TOLERANCE = 14


def _doy(month: int, day: int) -> int:
    """Day-of-year on a fixed non-leap calendar (for circular comparisons)."""
    return date(2001, month, day).timetuple().tm_yday


class YearLabeller:
    """Learns (date -> label) for one year column from existing template rows."""

    def __init__(self, pairs, tax_year: bool = False):
        """pairs – list of (date_or_datetime, label_value) from existing rows.
        tax_year – the column is a Self Assessment (SA) label: the UK tax
        year always starts on 6 April, so that boundary is used whenever the
        client's own rows agree with it."""
        self.groups = {}   # label_str -> {'label': original, 'min': date, 'max': date}
        for d, label in pairs:
            if d is None or label is None:
                continue
            if isinstance(d, datetime):
                d = d.date()
            if not isinstance(d, date):
                continue
            key = str(label).strip()
            if not key:
                continue
            g = self.groups.setdefault(key, {'label': label, 'min': d, 'max': d})
            if d < g['min']:
                g['min'] = d
            if d > g['max']:
                g['max'] = d

        # Cap every group's effective range at one year — anything beyond
        # the cap is a mislabelled straggler.
        for g in self.groups.values():
            g['cap'] = min(g['max'], g['min'] + _MAX_GROUP_SPAN)

        # Anchor: the most recent year group = the one that STARTS latest.
        self.anchor = None
        self.boundary = None       # (month, day) fiscal year start
        if self.groups:
            self.anchor = max(self.groups.values(), key=lambda g: g['min'])

            # Fiscal-year start day: the anchor group's first date, refined by
            # other groups whose first date lands near the same anniversary
            # (e.g. one group starts Dec 2, an older one Dec 1 -> use Dec 1).
            # Groups that start far from the anchor's anniversary began
            # mid-year (partial data) and are ignored.
            a_doy = _doy(self.anchor['min'].month, self.anchor['min'].day)
            candidates = [self.anchor['min']]
            for g in self.groups.values():
                if g is self.anchor:
                    continue
                g_doy = _doy(g['min'].month, g['min'].day)
                dist = abs(g_doy - a_doy)
                if min(dist, 365 - dist) <= _BOUNDARY_TOLERANCE:
                    candidates.append(g['min'])
            def rel(m):
                """Signed circular distance (days) from the anchor's start day."""
                return (_doy(m.month, m.day) - a_doy + 182) % 365 - 182

            best = min(candidates, key=rel)   # earliest observed start day
            self.boundary = (best.month, best.day)

            # Better: the year boundary must fall in the gap between one
            # year's last row and the next year's first row. A year's first
            # transaction is often weeks after its real start (Ahmed: 24/25
            # starts 6 Jul, but the year ends 31 May), so the gaps narrow it
            # down far more reliably than a first date does.
            refined = self._boundary_from_gaps()
            if refined:
                self.boundary = refined

            # An SA column with too little history to learn from (e.g. one
            # year, SA25 only) fell back to its first row's date — 3 May for
            # Kristaps, which labelled 7 Apr – 5 May 2025 as SA25 instead of
            # SA26. The tax year is fixed by law: use 6 April unless the
            # client's own rows contradict it.
            if tax_year and self._fits_boundary((4, 6)):
                self.boundary = (4, 6)

    def _fits_boundary(self, boundary) -> bool:
        """True when, with this year start, every label's rows fall in one
        year and different labels fall in different years."""
        saved, self.boundary = self.boundary, boundary
        try:
            seen = {}
            for key, g in self.groups.items():
                y0, y1 = self._year_index(g['min']), self._year_index(g['cap'])
                if y0 != y1 or seen.setdefault(y0, key) != key:
                    return False
            return True
        finally:
            self.boundary = saved

    def _boundary_from_gaps(self):
        """(month, day) that falls in the most year-to-year gaps, preferring
        6 April (UK tax year), then the 1st of a month. None if no gaps."""
        ordered = sorted(self.groups.values(), key=lambda g: g['min'])
        gaps = []      # (last date of one year, first date of the next]
        for prev, nxt in zip(ordered, ordered[1:]):
            if prev['cap'] < nxt['min']:
                gaps.append((prev['cap'], nxt['min']))
        if not gaps:
            return None

        def in_gap(m, d, lo, hi):
            for y in (hi.year - 1, hi.year):
                try:
                    b = date(y, m, d)
                except ValueError:
                    continue
                if lo < b <= hi:
                    return True
            return False

        days = [(m, d) for m in range(1, 13) for d in range(1, 32)
                if d <= (28 if m == 2 else 30 if m in (4, 6, 9, 11) else 31)]
        score = {md: sum(in_gap(*md, lo, hi) for lo, hi in gaps) for md in days}
        top = max(score.values())
        if top == 0:
            return None
        best = [md for md in days if score[md] == top]
        if (4, 6) in best:
            return (4, 6)
        firsts = [md for md in best if md[1] == 1]
        if not firsts:
            return None
        # several month starts fit (one wide gap): take the first one after
        # the previous year's last row, i.e. the year ended in that month
        lo = gaps[-1][0]
        return min(firsts, key=lambda md: (_doy(*md) - _doy(lo.month, lo.day)) % 365)

    @property
    def learned(self) -> bool:
        return self.anchor is not None

    def _year_index(self, d) -> int:
        """Which fiscal year (by start-calendar-year) does d fall in?"""
        if (d.month, d.day) >= self.boundary:
            return d.year
        return d.year - 1

    def label_for(self, d):
        """Return the label (same type/format as template) for date d."""
        if isinstance(d, datetime):
            d = d.date()
        if not self.learned:
            return None

        # Inside a known group's (straggler-capped) range?  When ranges
        # overlap, trust the group that starts latest.
        hits = [g for g in self.groups.values() if g['min'] <= d <= g['cap']]
        if hits:
            return max(hits, key=lambda g: g['min'])['label']

        # Extrapolate from the anchor group by fiscal-year index
        offset = self._year_index(d) - self._year_index(self.anchor['min'])
        return _shift_label(self.anchor['label'], offset)


def _shift_label(label, offset: int):
    """Add offset years to every numeric token in the label, preserving
    format and type.  'FY24'+1 -> 'FY25', 'SA24/25'+1 -> 'SA25/26',
    '2025'+1 -> '2026', 2025+1 -> 2026, 'SA24.25'+1 -> 'SA25.26'."""
    if offset == 0:
        return label

    if isinstance(label, (int, float)):
        return int(label) + offset

    s = str(label)

    def bump(match):
        tok = match.group()
        val = int(tok) + offset
        return str(val).zfill(len(tok))

    return re.sub(r'\d+', bump, s)


def learn_year_columns(ws, year_cols: dict, date_col: int, last_data_row: int) -> dict:
    """Build a YearLabeller for each year column from the template's rows.

    year_cols     – {'sa': col_idx_or_None, 'acc': ..., 'fy': ...}
    date_col      – 1-based column index of the Date column
    last_data_row – last row containing existing client data

    Returns {'sa': YearLabeller|None, ...}
    """
    out = {}
    for name, col in year_cols.items():
        if not col or not date_col:
            out[name] = None
            continue
        pairs = []
        for r in range(2, last_data_row + 1):
            d = ws.cell(r, date_col).value
            label = ws.cell(r, col).value
            if d is not None and label is not None:
                if isinstance(d, str):
                    d = _parse_date_str(d)
                if d is not None:
                    pairs.append((d, label))
        lab = YearLabeller(pairs, tax_year=(name == 'sa'))
        out[name] = lab if lab.learned else None
    return out


def _parse_date_str(s: str):
    s = str(s).strip()
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d %b %Y', '%d/%m/%y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None
