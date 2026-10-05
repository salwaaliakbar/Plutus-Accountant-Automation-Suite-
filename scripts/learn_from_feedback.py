"""Build core/learned_examples.json from client-corrected feedback workbooks.

The accountant reviews an output file and types the CORRECT category in the
'Correct category' column next to the AI's category column ('UC', 'UC
category', ...). Rows left blank there are skipped. This script turns those
corrections into (payee -> category) rules that every later run reuses.

Older feedback files with no 'Correct category' header are still read: the
column immediately right of the AI's category is used when its header is
empty. Formulas are never learned.

Run:  learn.bat   (or: python scripts/learn_from_feedback.py [Feedbacks-folder])
No restart needed — the backend reads the rules fresh on every upload.
"""

import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from core.categorise import _payee_key  # noqa: E402
from core.excel_writer import (  # noqa: E402
    CORRECTION_COLS, DESC_COLS, UC_COLS, _col, _read_headers,
)

# A payee's corrections must agree this often to become a rule.
MIN_SHARE = 0.6


def _correction_col(ws, hdrs: dict, uc_col: int):
    col = _col(hdrs, *CORRECTION_COLS)
    if col:
        return col
    legacy = uc_col + 1          # older files: unlabelled column next to it
    if legacy <= ws.max_column and ws.cell(1, legacy).value in (None, ''):
        return legacy
    return None


def learn(folder: Path) -> tuple[dict, list]:
    votes: dict[str, Counter] = defaultdict(Counter)
    sample_desc: dict[str, str] = {}
    report = []

    for path in sorted(folder.glob('*.xlsx')):
        if path.name.startswith('~$'):          # Excel's lock file
            continue
        wb = openpyxl.load_workbook(path, data_only=False)
        for sheet in wb.sheetnames:
            if 'raw' not in sheet.lower() or 'analysis' in sheet.lower():
                continue
            ws = wb[sheet]
            hdrs = _read_headers(ws)
            uc_col = _col(hdrs, *UC_COLS)
            desc_col = _col(hdrs, *DESC_COLS)
            corr_col = _correction_col(ws, hdrs, uc_col) if uc_col else None
            if not (uc_col and desc_col and corr_col):
                report.append(f"  skipped  {path.name} / {sheet}: no AI category, "
                              f"description or '{CORRECTION_COLS[0]}' column")
                continue

            n = 0
            for r in range(2, ws.max_row + 1):
                desc = ws.cell(r, desc_col).value
                corr = ws.cell(r, corr_col).value
                if not desc or not str(desc).strip() or corr is None:
                    continue
                corr = str(corr).strip()
                if not corr or corr.startswith('='):   # blank / formulas
                    continue
                corr = corr.rstrip('?').strip()        # '?' = accountant unsure
                key = _payee_key(desc)
                if not corr or not key:
                    continue
                votes[key][corr] += 1
                sample_desc.setdefault(key, str(desc)[:70])
                n += 1
            report.append(f"  read     {path.name} / {sheet}: {n} correction(s)")

    # Majority vote per payee; drop unclear ones
    examples = {}
    for key, counter in votes.items():
        top_cat, top_n = counter.most_common(1)[0]
        total = sum(counter.values())
        if top_n / total >= MIN_SHARE:
            examples[key] = {'category': top_cat, 'seen': total,
                             'example': sample_desc[key]}
    return examples, report


if __name__ == '__main__':
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / 'Feedbacks'
    out = ROOT / 'core' / 'learned_examples.json'
    before = {}
    if out.exists():
        before = json.loads(out.read_text(encoding='utf-8'))

    examples, report = learn(folder)
    print(f'Feedback files in {folder}:')
    print('\n'.join(report) or '  (none found)')

    with open(out, 'w', encoding='utf-8') as f:
        json.dump(examples, f, indent=1, ensure_ascii=False)

    new = {k: v for k, v in examples.items()
           if before.get(k, {}).get('category') != v['category']}
    print(f'\nLearned {len(examples)} payee rules ({len(new)} new or changed) -> {out}')
    for k, v in sorted(new.items(), key=lambda kv: -kv[1]['seen'])[:40]:
        print(f"  {v['seen']:3}x {k[:35]:<37} -> {v['category']}")
