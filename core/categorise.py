"""Transaction categorisation — the accountant's own decisions first, AI for the rest.

There are no hand-written keyword rules. Each transaction is categorised from:
 1. The client's template history: when the same payee (same direction of
    money) was always given one category in the client's existing rows,
    that category is reused directly — no AI call, same answer every run.
 2. core/learned_examples.json: corrections the accountant made on previous
    outputs (built by scripts/learn_from_feedback.py), reused the same way.
 3. Claude, for everything else, using ALL available columns (date,
    description, type, reference, bank category hint, amounts) with the
    client's rows as examples. Claude also says whether it is sure; rows it
    is not sure about are marked for the accountant to check.

Afterwards every category is reduced to one spelling per meaning ('DLA' vs
'Directors Loan Account'), preferring the client's own spelling, and
payments to the same person for the same purpose (e.g. several wordings of
a director's loan) are given one consistent category.

Privacy: account numbers, sort codes and IBANs are scrubbed from all text
before it is sent to the API.
"""

import json
import os
import re
from difflib import SequenceMatcher
from pathlib import Path

import anthropic
from dotenv import load_dotenv

from parsers._common import salvage_json_array

load_dotenv()

_MODEL = os.environ.get('CATEGORISE_MODEL', 'claude-sonnet-5')

# Master vocabulary: the client's supplied list plus categories the
# accountant actually used in corrected feedback files.
CATEGORIES = [
    'AOP', 'Accommodation', 'Accountancy', 'Accountancy fee', 'Bank charges',
    'Car', 'Car insurance', 'Car lease', 'Charging Car', 'Cleaning',
    'College', 'Company Car', 'Computer', 'CPD Grant', 'DBS', 'Dinner',
    'Directors Loan Account', 'DLA', 'Directors salary', 'Dividends',
    'Donation', 'Education Course', 'Entertainment', 'Equipment', 'FODO',
    'Food', 'GOC', 'Gym', 'Heat', 'HMRC', 'HMRC-CT', 'HMRC-PAYE', 'HMRC-SA',
    'HMRC-VAT', 'HMRC - maternity pay', 'Hotel', 'Income', 'Insurance',
    'Insurance - AOP', 'Interest income', 'Internet', 'Light and heat',
    'Lunch', 'Marketing', 'Mobile phone', 'Other direct costs', 'Parking',
    'PCSE', 'Penalty fee', 'Petrol', 'Postage', 'Professional fee',
    'Professional - College', 'Professional - AOP', 'Professional - GOC',
    'Purchase', 'Refund', 'Rent', 'Repairs', 'Sales', 'Security', 'SIPP',
    'SMP', 'Staff benefits', 'Stationary', 'Subscriptions', 'Sundry',
    'Taxi', 'Telephone', 'Toll', 'Train', 'Transfer', 'Travel',
    'Trivial benefits', 'Wages', 'Wages-staff', 'Water', 'Website',
    'Work from home', 'Unknown',
]

_LEARNED_PATH = Path(__file__).parent / 'learned_examples.json'

# Spellings that mean the same category. The first entry is the default;
# when the client already uses one of the spellings, theirs wins.
SYNONYMS = [
    ['DLA', 'Directors Loan Account', "Director's Loan Account", 'Directors loan',
     "Director's loan", 'Director loan'],
    ['Directors salary', "Director's salary", 'Director salary'],
    ['Wages-staff', 'Wages - staff', 'Staff wages', 'Salary-Staff', 'Staff salary'],
    ['Accountancy', 'Accountancy fee', 'Accountant fee', 'Accounting fee'],
    ['Bank charges', 'Bank charge', 'Bank charge?', 'Bank fee'],
    ['Mobile', 'Mobile phone'],
    ['Professional fee', 'Professional fees'],
    ['Stationary', 'Stationery'],
    ['Subscriptions', 'Subscription'],
    ['Dividends', 'Dividend'],
    ['Sundry', 'Sundries'],
    ['Refund', 'Refunds'],
    # professional bodies: a client who books these as plain 'College' /
    # 'AOP' / 'GOC' must not get a second 'Professional - …' row for them
    ['College', 'Professional - College', 'College of Optometrists'],
    ['AOP', 'Professional - AOP', 'Insurance - AOP'],
    ['GOC', 'Professional - GOC'],
    ['Donation', 'Donations', 'Charity'],
]

# How sure the client's history must be before its category is reused
# without asking the AI (share of that payee's past rows).
_HISTORY_MIN_SHARE = 0.75


# ── Privacy scrub ─────────────────────────────────────────────────────────────

_SENSITIVE_RE = re.compile(
    r'\bA/?C\s*:?\s*\d{6,}\b'          # "A/C 41204042"
    r'|\b\d{2}-\d{2}-\d{2}\b'          # sort code
    r'|\b\d{8,}\b'                     # bare 8+ digit account-like numbers
    r'|\bIBAN\s*:?\s*[A-Z]{2}\d{2}[A-Z0-9]{4,}\b',
    re.IGNORECASE,
)


def _scrub(text) -> str:
    return _SENSITIVE_RE.sub('#', str(text or ''))


# ── Payee normalisation (shared with learn_from_feedback) ─────────────────────

_STOP = {'CARD', 'PAYMENT', 'TO', 'ON', 'DIRECT', 'DEBIT', 'CREDIT', 'AUTOMATED',
         'ONLINE', 'TRANSACTION', 'VIA', 'MOBILE', 'FP', 'THE', 'REF', 'BGC',
         'TFR', 'DD', 'SO', 'FASTER', 'INWARD', 'OUTWARD', 'BACS', 'RECEIVED',
         'TRANSFER', 'PYMT', 'IN', 'OUT', 'LVP'}


def _payee_key(desc: str) -> str:
    s = str(desc).upper()
    s = re.sub(r'\d+', ' ', s)
    s = re.sub(r'[^A-Z& ]+', ' ', s)
    words = [w for w in s.split() if len(w) > 1 and w not in _STOP]
    return ' '.join(words[:4])


def _load_learned() -> dict:
    if _LEARNED_PATH.exists():
        try:
            with open(_LEARNED_PATH, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


# ── Example assembly ──────────────────────────────────────────────────────────

def _build_examples(client_examples: list | None) -> tuple[list, list]:
    """Merge template examples (priority) with global learned corrections.

    client_examples – list of dicts {'description','type','reference','category'}
                       extracted from the client's own template rows.
    Returns (example_lines, client_categories).
    """
    seen_keys = {}
    lines = []
    client_cats = []

    # 1. Client's own template rows (authoritative for THIS client)
    for ex in (client_examples or []):
        cat = str(ex.get('category') or '').strip()
        desc = str(ex.get('description') or '').strip()
        if not cat or not desc or cat.startswith('='):
            continue
        if cat not in client_cats:
            client_cats.append(cat)
        key = _payee_key(desc)
        if not key or key in seen_keys:
            continue
        seen_keys[key] = True
        ref = str(ex.get('reference') or '').strip()
        typ = str(ex.get('type') or '').strip()
        extra = f' | ref: {_scrub(ref)}' if ref else ''
        extra += f' | type: {typ}' if typ else ''
        lines.append(f'"{_scrub(desc)[:60]}"{extra} -> {cat}')

    # 2. Global corrections from past accountant feedback
    for key, info in _load_learned().items():
        if key in seen_keys:
            continue
        seen_keys[key] = True
        lines.append(f'"{_scrub(info["example"])[:60]}" -> {info["category"]}')

    return lines[:150], client_cats


# ── Category spelling: one name per meaning ───────────────────────────────────

def _cat_key(s) -> str:
    """'Bank charge?' / 'bank charges' / 'Bank-Charge' -> 'bankcharge'."""
    k = re.sub(r'[^a-z0-9]+', '', str(s or '').lower())
    return k[:-1] if len(k) > 3 and k.endswith('s') else k


def _make_canonicaliser(client_cats: list, analysis_vocab: list | None = None):
    """Return a function mapping any category spelling to the one to use:
    the client's own spelling when they have one for that meaning (their
    RAW rows first, then the names their Analysis formulas look up), else
    the group's default, else the master list's spelling, else the first
    spelling seen this run. Analysis names also match despite small typos
    ('Inusrance - AOP' for 'Insurance - AOP')."""
    client_by_key = {_cat_key(c): c for c in analysis_vocab or []}
    client_by_key.update({_cat_key(c): c for c in client_cats})
    fuzzy_pool = [(k, c) for k, c in ((_cat_key(c), c) for c in analysis_vocab or [])
                  if len(k) >= 6]

    def client_spelling(k):
        if k in client_by_key:
            return client_by_key[k]
        if len(k) >= 6:
            for ck, c in fuzzy_pool:
                if SequenceMatcher(None, k, ck).ratio() >= 0.9:
                    return c
        return None

    master_by_key = {_cat_key(c): c for c in CATEGORIES}
    group_of = {}
    for group in SYNONYMS:
        keys = [_cat_key(s) for s in group]
        raw_hit = next((c for c in client_cats if _cat_key(c) in keys), None)
        chosen = raw_hit or next(
            (s for s in (client_spelling(k) for k in keys) if s), group[0])
        for k in keys:
            group_of[k] = chosen
    seen = {}

    def canon(cat) -> str:
        cat = str(cat or '').strip()
        if not cat:
            return 'Unknown'
        k = _cat_key(cat)
        if k in client_by_key:
            return client_by_key[k]
        if k in group_of:
            return group_of[k]
        mine = client_spelling(k)       # small typos in the Analysis names
        if mine:
            return mine
        if k in master_by_key:
            return master_by_key[k]
        return seen.setdefault(k, cat)

    return canon


# ── Reuse the accountant's own past decisions ─────────────────────────────────

def _direction(t: dict) -> str:
    return 'in' if t.get('money_in', 0) > 0 else 'out'


def _history_rules(client_examples: list | None) -> tuple[dict, set]:
    """Payee -> category from the client's own rows, only where the history
    is clear. Returns ({(payee_key, direction): category}, ambiguous_keys).
    Rows whose direction is unknown are stored under direction None."""
    votes: dict[tuple, dict] = {}
    for ex in client_examples or []:
        cat = str(ex.get('category') or '').strip()
        key = _payee_key(ex.get('description') or '')
        if not cat or cat.startswith('=') or len(key) < 3:
            continue
        slot = votes.setdefault((key, ex.get('direction')), {})
        slot[cat] = slot.get(cat, 0) + 1

    rules, ambiguous = {}, set()
    for k, counts in votes.items():
        top = max(counts, key=counts.get)
        if counts[top] / sum(counts.values()) >= _HISTORY_MIN_SHARE:
            rules[k] = top
        else:
            ambiguous.add(k[0])
    return rules, ambiguous


def _rule_for(t: dict, rules: dict, ambiguous: set, learned: dict):
    """(category, source) from past decisions, or (None, None) to ask the AI."""
    key = _payee_key(t.get('description') or '')
    if len(key) < 3:
        return None, None
    d = _direction(t)
    for k in ((key, d), (key, None)):
        if k in rules:
            return rules[k], 'history'
    # the client's history disagrees with itself, or only has this payee in
    # the other direction: their own rows beat corrections from other files
    if key in ambiguous or (key, 'in' if d == 'out' else 'out') in rules:
        return None, None
    info = learned.get(key)
    if info and info.get('category'):
        return info['category'], 'learned'
    return None, None


# ── Payments to the same person for the same purpose ──────────────────────────

_PURPOSES = (
    ('loan',     re.compile(r'\bLOAN\b|\bDLA\b', re.I)),
    ('dividend', re.compile(r'DIVID', re.I)),
    ('salary',   re.compile(r'SALARY|\bWAGES?\b', re.I)),
)
_COMPANY_RE = re.compile(r'\b(LTD|LIMITED|PLC|LLP|BANK|CREDIT|FINANCE|SERVICES)\b', re.I)
_NAME_STOP = {'TO', 'AC', 'VIA', 'MOBILE', 'XFER', 'PYMT', 'FP', 'MR', 'MRS', 'MS', 'DR'}


def _person_tokens(desc: str) -> set:
    """Name words of the person paid, e.g. 'To A/C 80522920 , J P PATEL ,
    Via Mobile Xfer , Loan' -> {'PATEL'}. Empty for companies."""
    parts = [p.strip() for p in str(desc or '').split(',') if p.strip()]
    if parts and re.match(r'^TO\s+A/?C\b', parts[0], re.I):
        parts = parts[1:]
    if not parts or _COMPANY_RE.search(parts[0]):
        return set()
    words = re.sub(r'[^A-Z ]+', ' ', parts[0].upper()).split()
    return {w for w in words if len(w) >= 3 and w not in _NAME_STOP}


def _same_person(a: set, b: set) -> bool:
    for x in a:
        for y in b:
            if x == y or (min(len(x), len(y)) >= 4
                          and SequenceMatcher(None, x, y).ratio() >= 0.8):
                return True
    return False


def _harmonise_people(transactions, results, sources, client_cats) -> int:
    """Give payments out to the same person with the same purpose word
    (loan / dividend / salary) one category — the most common one among
    them. Only AI answers are changed; the accountant's own are kept.
    Returns how many rows changed."""
    groups = []    # [purpose, name tokens, [indices]]
    for i, t in enumerate(transactions):
        if _direction(t) != 'out':
            continue
        text = f"{t.get('description', '')} {t.get('reference', '')}"
        purpose = next((p for p, rx in _PURPOSES if rx.search(text)), None)
        tokens = _person_tokens(t.get('description', ''))
        if not purpose or not tokens:
            continue
        for g in groups:
            if g[0] == purpose and _same_person(g[1], tokens):
                g[1] |= tokens
                g[2].append(i)
                break
        else:
            groups.append([purpose, set(tokens), [i]])

    changed = 0
    for _, _, idx in groups:
        if len(idx) < 2:
            continue
        counts = {}
        for i in idx:
            if results[i] != 'Unknown':
                counts[results[i]] = counts.get(results[i], 0) + 1
        if not counts:
            continue
        # most common; ties go to the client's own vocabulary
        best = max(counts, key=lambda c: (counts[c], c in client_cats))
        for i in idx:
            if sources[i] == 'ai' and results[i] != best:
                results[i] = best
                changed += 1
    return changed


def _fill_unknown_from_same_payee(transactions, results) -> list[int]:
    """An 'Unknown' row whose payee (same direction of money) got exactly one
    category everywhere else in this statement takes that category — e.g.
    one 'OPTIC VILLAGE LTD' credit left Unknown while every other one is
    Income. Returns the indices changed (the caller keeps them marked for
    review)."""
    seen: dict[tuple, set] = {}
    for t, c in zip(transactions, results):
        key = _payee_key(t.get('description') or '')
        if len(key) >= 3 and c != 'Unknown':
            seen.setdefault((key, _direction(t)), set()).add(c)
    changed = []
    for i, t in enumerate(transactions):
        if results[i] != 'Unknown':
            continue
        cats = seen.get((_payee_key(t.get('description') or ''), _direction(t)))
        if cats and len(cats) == 1:
            results[i] = next(iter(cats))
            changed.append(i)
    return changed


# ── Claude call ───────────────────────────────────────────────────────────────

def _cache_key(t: dict) -> tuple:
    return (
        str(t.get('description', '')).upper(),
        str(t.get('reference', '')).upper(),
        _direction(t),
    )


def _build_system_prompt(example_lines: list, client_cats: list) -> str:
    cats = ', '.join(f"'{c}'" for c in CATEGORIES)
    client_block = ''
    if client_cats:
        client_block = (
            "\nThis client's own category vocabulary (STRONGLY prefer these exact "
            "spellings when they apply):\n" + ', '.join(f"'{c}'" for c in client_cats) + '\n'
        )
    ex_block = ''
    if example_lines:
        ex_block = (
            '\nExamples of how this accountant categorises (payee -> category):\n'
            + '\n'.join(example_lines) + '\n'
        )

    return f"""You are an expert UK accountant's assistant categorising bank transactions
for an optician / optometrist practice. For every transaction you receive the
date, description, transaction type, payment reference, the bank's own category
hint, and the money in/out amounts. Use ALL of these fields, not just the
description. The payment reference is often decisive — e.g. a transfer to the
director with reference 'Salary' is a salary, with reference 'Dividend' it is a dividend.
{client_block}{ex_block}
General category list (fallback when the client vocabulary doesn't cover it):
{cats}

Domain knowledge:
- AOP = Association of Optometrists (payee often 'ASSOCIATION OF OPTOMETRISTS');
  GOC = General Optical Council (payee often 'WWW.OPTICAL.ORG' or 'GENERAL OPTICAL COUNCIL' —
  this is their annual registration fee, category 'GOC', not 'Professional fee');
  FODO = opticians' trade body; PCSE = Primary Care Support England (NHS income);
  DBS = Disclosure and Barring Service; SIPP = personal pension contributions
  (e.g. Hargreaves Lansdown, Vanguard, AJ Bell).
- Payments to the business owner: use the reference and amount pattern to
  distinguish salary (regular, fixed, often ref 'Salary') from dividends
  (ref 'Dividend' or irregular round amounts) from DLA (Directors Loan Account).
- HMRC payments: use the reference to pick HMRC-PAYE / HMRC-VAT / HMRC-CT /
  HMRC-SA. Plain 'HMRC' only when the reference gives no clue.
- Credits from opticians chains (Specsavers, Vision Express, Boots Opticians),
  NHS/PCSE, or locum agencies are 'Income'.
- Petrol stations (Shell, BP, Esso, Texaco, Rontec, MFG, Sainsburys Petrol,
  garage names) are 'Petrol', not Equipment or Food.
- Supermarket small amounts near lunchtime, cafes, bakeries (Greggs, Caffe Nero,
  Sainsbury's Local small amounts) are usually 'Lunch'.

Rules:
1. Return ONLY a JSON array with one object per transaction, same order:
   {{"category": "<category>", "sure": true}}
   Set "sure" to false whenever an accountant should double-check the
   answer: the payee is unfamiliar, several categories are plausible, or
   you guessed.
2. Prefer the client's own vocabulary above; then the general list.
3. If genuinely nothing fits, invent the most appropriate short accounting
   category name (e.g. 'Software', 'Cleaning') — do NOT force a bad match.
4. Use 'Unknown' only when the transaction cannot be identified at all."""


def _batch_classify(transactions: list[dict], system_prompt: str) -> list[tuple[str, bool]]:
    """[(category, sure)] for each transaction, same order."""
    client = anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])

    items = []
    for t in transactions:
        direction = 'IN' if t.get('money_in', 0) > 0 else 'OUT'
        amount = t.get('money_in', 0) if direction == 'IN' else t.get('money_out', 0)
        d = t.get('date')
        parts = [
            f"date: {d}",
            f"desc: {_scrub(t.get('description', ''))}",
        ]
        if t.get('subcategory'):
            parts.append(f"type: {t['subcategory']}")
        if t.get('reference'):
            parts.append(f"ref: {_scrub(t['reference'])}")
        if t.get('csv_category'):
            parts.append(f"bank-hint: {t['csv_category']}")
        parts.append(f"{direction} {amount:.2f}")
        items.append(' | '.join(parts))

    user_msg = 'Categorise these transactions:\n' + '\n'.join(
        f"{i + 1}. {item}" for i, item in enumerate(items)
    )

    response = client.messages.create(
        model=_MODEL,
        max_tokens=8000,
        system=system_prompt,
        messages=[{'role': 'user', 'content': user_msg}],
    )

    text = ''.join(
        block.text for block in response.content if getattr(block, 'type', '') == 'text'
    ).strip()
    # Claude's extended-thinking budget is drawn from the same max_tokens
    # limit as the reply, so a big batch can get cut off mid-array before
    # any JSON closes. salvage_json_array keeps every complete leading
    # answer instead of raising and losing the whole batch (and with it,
    # the whole upload) over a truncated tail.
    answers = salvage_json_array(text)
    if not isinstance(answers, list) or not answers:
        # sometimes the answers come one JSON object per line, not as an array
        answers = []
        for m in re.finditer(r'\{[^{}]*\}', text):
            try:
                answers.append(json.loads(m.group()))
            except json.JSONDecodeError:
                pass
    if not answers:
        raise ValueError(f"Claude returned unexpected output: {text[:200]!r}")

    out = []
    for a in answers:
        if isinstance(a, dict):
            cat, sure = str(a.get('category') or '').strip(), a.get('sure') is not False
        else:                        # a bare string, the older answer format
            cat, sure = str(a).strip(), True
        out.append((cat or 'Unknown', sure))

    n = len(transactions)
    if len(out) > n:
        out = out[:n]
    elif len(out) < n:
        out += [('Unknown', False)] * (n - len(out))
    return out


def categorise_detailed(
    transactions: list[dict],
    client_examples: list | None = None,
    batch_size: int = 40,
    analysis_vocab: list | None = None,
) -> tuple[list[str], list[bool]]:
    """Return (categories, needs_review), one entry per transaction.

    needs_review is True where the AI was not sure or answered 'Unknown';
    answers taken from the accountant's own past decisions never are.
    client_examples – (description/type/reference/category/direction) dicts
    extracted from the client's template.
    analysis_vocab – category names the template's Analysis formulas look
    up (see excel_writer.extract_analysis_vocabulary); answers are spelled
    to match them.
    """
    example_lines, client_cats = _build_examples(client_examples)
    vocab = client_cats + [c for c in analysis_vocab or []
                           if _cat_key(c) not in {_cat_key(x) for x in client_cats}]
    system_prompt = _build_system_prompt(example_lines, vocab)
    canon = _make_canonicaliser(client_cats, analysis_vocab)
    rules, ambiguous = _history_rules(client_examples)
    learned = _load_learned()

    n = len(transactions)
    results: list = [None] * n
    sure: list = [True] * n
    sources: list = ['ai'] * n

    # 1. the accountant's own past decisions
    for i, t in enumerate(transactions):
        cat, src = _rule_for(t, rules, ambiguous, learned)
        if cat:
            results[i], sources[i] = cat, src
    n_hist, n_learn = sources.count('history'), sources.count('learned')
    print(f"  [categorise] {n_hist} from client history, {n_learn} from learned "
          f"corrections, {n - n_hist - n_learn} left for the AI")

    # 2. the AI for the rest. Identical transactions are asked about once, so
    # duplicates split across batches can't get different answers.
    key_to_indices: dict[tuple, list[int]] = {}
    for i, t in enumerate(transactions):
        if results[i] is None:
            key_to_indices.setdefault(_cache_key(t), []).append(i)
    unique_keys = list(key_to_indices.keys())
    todo_txn = [transactions[key_to_indices[k][0]] for k in unique_keys]

    if todo_txn:
        api_key = os.environ.get('ANTHROPIC_API_KEY', '')
        if not api_key or api_key.startswith('sk-ant-...'):
            print(f"  [categorise] No API key — marking {len(todo_txn)} txns as 'Unknown'")
            answers = [('Unknown', False)] * len(todo_txn)
        else:
            print(f"  [categorise] Model: {_MODEL}, examples: {len(example_lines)} "
                  f"({len(client_cats)} client categories), "
                  f"{len(todo_txn)} unique transactions")
            answers = []
            for start in range(0, len(todo_txn), batch_size):
                batch = todo_txn[start:start + batch_size]
                print(f"  [categorise] Claude batch {start // batch_size + 1}: {len(batch)} txns...")
                answers += _batch_classify(batch, system_prompt)
        for k, (cat, is_sure) in zip(unique_keys, answers):
            for i in key_to_indices[k]:
                results[i], sure[i] = cat, is_sure

    # 3. one spelling per meaning, then one answer per person + purpose
    results = [canon(c) for c in results]
    changed = _harmonise_people(transactions, results, sources, client_cats)
    if changed:
        print(f"  [categorise] {changed} payment(s) to the same person aligned")
    filled = _fill_unknown_from_same_payee(transactions, results)
    for i in filled:
        sure[i] = False             # still worth a look
    if filled:
        print(f"  [categorise] {len(filled)} 'Unknown' row(s) given their payee's category")

    review = [src == 'ai' and (not s or c == 'Unknown')
              for c, s, src in zip(results, sure, sources)]
    return results, review


def categorise(
    transactions: list[dict],
    client_examples: list | None = None,
    batch_size: int = 40,
) -> list[str]:
    """Return a category for each transaction (see categorise_detailed)."""
    return categorise_detailed(transactions, client_examples, batch_size)[0]
