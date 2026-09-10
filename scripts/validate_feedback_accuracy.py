"""Held-out accuracy check against the client corrections in Feedbacks/.

Unlike scripts/validate_accuracy.py (which uses a client's OWN historic
workbook as ground truth), this validates against the corrections the
accountant wrote into the Feedbacks/*.xlsx review files: the column to the
right of 'UC category' / 'UC Category'.

For each file, the OLDER rows are used to build the learned-example pool
(simulating "corrections already absorbed before this point") and the most
recent chunk is held out as TEST — the real categorise() pipeline (current
code + the rebuilt learned_examples.json) is run on the test rows and
compared against what the accountant actually wrote.

Run:  venv\\Scripts\\python scripts\\validate_feedback_accuracy.py
"""

import re
import sys
from pathlib import Path

import openpyxl
from dotenv import load_dotenv

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
load_dotenv()

from core import categorise as categorise_mod  # noqa: E402
from core.categorise import categorise  # noqa: E402
from scripts.learn_from_feedback import _payee_key  # noqa: E402

TEST_FRACTION = 0.25

FILES = {
    'sample1 (Matthew Ferris)': dict(
        fname='output test sample 1.xlsx', sheet='RAW 2025',
        date=4, desc=5, ref=6, type=7, csv_cat=15, in_=9, out=10, uc=16, corr=17,
    ),
    'sample2 (Chido Hove)': dict(
        fname='output test sample 2.xlsx', sheet='Raw 24',
        date=4, desc=5, ref=None, type=None, csv_cat=None, in_=6, out=7, uc=12, corr=13,
    ),
    'sample5 (Ahmed Ibrahim)': dict(
        fname='output test sample 5.xlsx', sheet='RAW25',
        date=3, desc=4, ref=None, type=5, csv_cat=None, in_=6, out=7, uc=10, corr=11,
    ),
}


def _cell(ws, r, c):
    return ws.cell(r, c).value if c else None


def load_rows(cols):
    wb = openpyxl.load_workbook(ROOT / 'Feedbacks' / cols['fname'], data_only=False)
    ws = wb[cols['sheet']]
    rows = []
    for r in range(2, ws.max_row + 1):
        desc = _cell(ws, r, cols['desc'])
        uc = _cell(ws, r, cols['uc'])
        corr = _cell(ws, r, cols['corr'])
        if not desc or corr is None or not str(corr).strip():
            continue
        corr_s = str(corr).strip()
        if corr_s.startswith('='):
            continue  # broken formula, not real feedback
        if corr_s.endswith('?'):
            continue  # accountant themself was unsure -- not gradeable ground truth
        money_in = _cell(ws, r, cols['in_']) or 0
        money_out = _cell(ws, r, cols['out']) or 0
        try:
            money_in = float(money_in)
        except (TypeError, ValueError):
            money_in = 0.0
        try:
            money_out = float(money_out)
        except (TypeError, ValueError):
            money_out = 0.0
        rows.append({
            'date': _cell(ws, r, cols['date']),
            'description': str(desc).strip(),
            'reference': str(_cell(ws, r, cols['ref']) or '').strip(),
            'subcategory': str(_cell(ws, r, cols['type']) or '').strip(),
            'csv_category': str(_cell(ws, r, cols['csv_cat']) or '').strip(),
            'money_in': money_in,
            'money_out': money_out,
            'uc_category': str(uc).strip() if uc else '',
            'category': corr_s,
        })
    return rows


def norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower())


def build_learned_pool(train_rows_by_file):
    """Same majority-vote logic as learn_from_feedback.learn(), but over
    in-memory TRAIN rows only (no test-set leakage)."""
    from collections import Counter, defaultdict
    votes = defaultdict(Counter)
    sample_desc = {}
    for rows in train_rows_by_file:
        for r in rows:
            key = _payee_key(r['description'])
            if not key:
                continue
            votes[key][r['category']] += 1
            sample_desc.setdefault(key, r['description'][:70])
    examples = {}
    for key, counter in votes.items():
        (top_cat, top_n), *_ = counter.most_common()
        total = sum(counter.values())
        if top_n / total >= 0.6:
            examples[key] = {'category': top_cat, 'seen': total, 'example': sample_desc[key]}
    return examples


def run(name, train, test, learned_pool):
    client_examples = [
        {'description': r['description'], 'type': r['subcategory'],
         'reference': r['reference'], 'category': r['category']}
        for r in train
    ]
    test_txns = [
        {k: r[k] for k in ('date', 'description', 'reference', 'subcategory',
                            'csv_category', 'money_in', 'money_out')}
        for r in test
    ]
    truth = [r['category'] for r in test]
    baseline = [r['uc_category'] for r in test]

    categorise_mod._load_learned = lambda: learned_pool
    categorise_mod._CACHE.clear()

    print(f"\n=== {name} === train={len(train)} test={len(test)}")
    preds = categorise(test_txns, client_examples=client_examples)

    def score(preds):
        exact = fuzzy = 0
        mism = []
        for t, p, txn in zip(truth, preds, test):
            if p == t:
                exact += 1; fuzzy += 1
            elif norm(p) == norm(t):
                fuzzy += 1
            else:
                mism.append((txn['description'][:45], p, t))
        return exact, fuzzy, mism

    b_exact, b_fuzzy, _ = score(baseline)
    n = len(test)
    print(f"  ORIGINAL output accuracy on this held-out slice: {100*b_exact/n:.1f}% exact / {100*b_fuzzy/n:.1f}% fuzzy")

    exact, fuzzy, mismatches = score(preds)
    print(f"  NEW pipeline accuracy (after fixes):             {100*exact/n:.1f}% exact / {100*fuzzy/n:.1f}% fuzzy")
    if mismatches:
        print("  Remaining mismatches (desc | new AI -> correct):")
        for desc, p, t in mismatches:
            print(f"    \"{desc}\" | {p!r} -> {t!r}")
    return n, exact, fuzzy, b_exact, b_fuzzy


if __name__ == '__main__':
    all_rows = {name: load_rows(cols) for name, cols in FILES.items()}

    splits = {}
    for name, rows in all_rows.items():
        cut = int(len(rows) * (1 - TEST_FRACTION))
        splits[name] = (rows[:cut], rows[cut:])

    learned_pool = build_learned_pool([tr for tr, _ in splits.values()])
    print(f"Built learned pool from TRAIN rows only: {len(learned_pool)} payee patterns "
          f"(test rows never seen by the learner)")

    totals = {'n': 0, 'exact': 0, 'fuzzy': 0, 'b_exact': 0, 'b_fuzzy': 0}
    for name, (train, test) in splits.items():
        if len(test) < 5:
            print(f"\n=== {name} === too few test rows, skipping")
            continue
        n, exact, fuzzy, b_exact, b_fuzzy = run(name, train, test, learned_pool)
        totals['n'] += n
        totals['exact'] += exact
        totals['fuzzy'] += fuzzy
        totals['b_exact'] += b_exact
        totals['b_fuzzy'] += b_fuzzy

    print("\n=== OVERALL (held-out test rows only) ===")
    n = totals['n']
    print(f"  ORIGINAL: {totals['b_exact']}/{n} = {100*totals['b_exact']/n:.1f}% exact, "
          f"{100*totals['b_fuzzy']/n:.1f}% fuzzy")
    print(f"  NEW:      {totals['exact']}/{n} = {100*totals['exact']/n:.1f}% exact, "
          f"{100*totals['fuzzy']/n:.1f}% fuzzy")
