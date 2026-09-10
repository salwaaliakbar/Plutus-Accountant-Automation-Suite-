# Plutus — Client Guide
### How the Bank Statement Processor works (plain English)

---

## Your most important questions answered

---

### "Was the Excel file brand new or an existing one?"

**It is your existing Excel workbook** — the one called `Bank summarised Behesta v1 2025.xlsx`.

That file already had years of historical transactions in it (1,519 rows of old data going back to 2021). The system does **not** create a new file from scratch and does **not** touch any of your old data.

What it does is simple:

```
BEFORE:
  Your Excel workbook
  Row 1:    Header (No, SA, ACC, Date, Subcategory, Memo, Paid in, Paid out...)
  Row 2:    Sep 2022 transaction
  Row 3:    Sep 2022 transaction
  ...
  Row 1520: Last existing transaction (your old data ends here)

AFTER uploading May 2024 statement:
  Row 1:    Header  ← untouched
  Row 2:    Sep 2022 transaction  ← untouched
  ...
  Row 1520: Last existing transaction  ← untouched
  Row 1521: ← NEW → 02/05/2024  Income   Ayaoptics Ltd  £3,025.00
  Row 1522: ← NEW → 01/05/2024  Travel   TFL Travel     £7.90
  Row 1523: ← NEW → 01/05/2024  Lunch    SQ Blank St    £4.40
  ...
  Row 1567: ← NEW → 08/04/2024  Bank charges  Commission  £8.50
```

**All 47 new transactions from May 2024 are added at the bottom. Nothing else changes.**

---

### "What if it's a CSV file?"

CSV and PDF both work **exactly the same way** — you get the same result either way.

Barclays lets you download your statement in different formats. This system accepts all of them:

| Format | How to download from Barclays | Result |
|--------|-------------------------------|--------|
| **CSV** | Online banking → Statements → Download → CSV | Same 47 rows appended |
| **PDF** | Online banking → Statements → Download → PDF | Same 47 rows appended |
| **XLSX** | Online banking → Statements → Download → Excel | Same 47 rows appended |

The system reads whichever format you give it and produces the same output — your master Excel workbook with the new month's transactions added.

---

### "What did the new bit do exactly?"

For each new transaction the system automatically:

1. **Reads the date** and works out the UK financial year
   - e.g. 2 May 2024 → `FY24`, `24/25`

2. **Recognises the payee** and assigns a category
   - TFL → *Travel*
   - Blank Street / Joe the Juice → *Lunch*
   - Ayaoptics / Specsavers → *Income*
   - Miss B Hamid salary → *Directors salary*
   - etc.

3. **Writes the row** into your RAW tab with the correct columns and formulas:
   - Column J (NET) = Paid in minus Paid out (live formula)
   - Column K (Balance UC) = running total that chains from the previous row (live formula)
   - Column L (Check UC) = compares the printed balance against the running total (live formula)

4. **Rebuilds the Analysis pivot** so the summary tab automatically includes the new month's figures — no manual work needed.

---

### "What does the downloaded file contain?"

When you click **Download Excel**, you receive a copy of your master workbook with the new month already appended. It is a standard `.xlsx` file you can open in Excel immediately.

You will see:
- **RAW (2) tab** — all your historical rows plus the new month's rows at the bottom
- **Analysis 24 Raw (2) tab** — the summary pivot, automatically updated to include the new data

---

### "How does the category assignment work?"

Every transaction is categorised by **Claude AI** — there is no fixed rule list anymore. The AI is shown, for every transaction, the full picture: date, description, transaction type, payment reference, the bank's own category hint, and whether money went in or out. It doesn't just pattern-match a payee name — it reasons about the whole transaction, which is why it can correctly tell a salary payment apart from a dividend or a director's loan repayment to the same person, just from the reference and amount.

Before categorising, the AI is given three things so it categorises *your* business the way *your* accountant does:

1. **Your own workbook's history** — the categories your accountant already assigned in your existing spreadsheet rows are read and given to the AI as the strongest guide. Your own spelling and conventions always come first.
2. **Corrections learned from past reviews** — see below.
3. **A general accounting category list** and optician/optometry-specific domain knowledge (e.g. recognising AOP, GOC, PCSE, FODO, DBS, SIPP, and how to distinguish salary vs. dividend vs. director's loan) as a fallback for anything your own history doesn't cover.

Only the description and reference text are sent to the AI, and account numbers, sort codes and IBANs are automatically stripped out first — never any full account details.

---

### "How does the system get smarter over time?"

Every time you review an output file and correct a category (writing the right one in the column next to "UC Category"), that correction isn't just fixing that one file — it's saved. When enough corrections agree on how a particular payee should be categorised, that becomes a permanent example the AI is shown for every future statement, for every client. So a mistake only needs to be corrected once — after that, the same payee is categorised correctly automatically.

---

### "What categories does it use?"

The AI prefers your own workbook's existing category names first. When nothing in your history fits, it falls back to a general accounting category list (Accountancy, Bank charges, HMRC-PAYE/VAT/CT/SA, Directors salary, DLA, Dividends, Income, Lunch, Petrol, Travel, Subscriptions, Sundry, and more) — and if genuinely nothing fits, it proposes the most sensible short category name rather than forcing a bad match.

If you ever want a category corrected for a specific payee, just write the right category next to it in your review file — that correction gets learned automatically (see above), no code changes needed.

---

### "What does the system NOT do?"

- It does **not** modify the original statement file (CSV/PDF/XLSX) you uploaded
- It does **not** delete or change any of your existing historical data
- It does **not** overwrite your master template — it makes a copy and appends to that copy
- It does **not** send your account number, sort code, IBAN or name to any external service

---

### Step-by-step: what happens when you upload a file

```
1. You drag your Barclays statement onto the website
         ↓
2. The system reads every transaction from the file
         ↓
3. Each transaction is checked against the rule list
   → Known payee? → Category assigned instantly
   → Unknown payee? → Sent to Claude AI for categorisation
         ↓
4. UK financial year labels are calculated for each transaction
         ↓
5. All transactions are written into your Excel workbook
   (appended after the last existing row, with live formulas)
         ↓
6. The Analysis pivot is rebuilt to include the new data
         ↓
7. A download button appears → click to save the updated workbook
```

Total time: typically **15–30 seconds** per statement.

---

### Where does the Excel template live?

The master template is stored in the `Samples` folder:

```
Samples/
└── Bank summarised Behesta v1 2025.xlsx   ← this is your template
```

The system always reads from this template and saves the result to the `output` folder. Your template is never modified.

```
output/
└── [processed file].xlsx   ← this is what you download
```

---

### Quick reference — running the system

Open two command prompt windows:

**Window 1:**
```
cd "Plutus Accountant Automation Suite"
venv\Scripts\uvicorn api.server:app --host 127.0.0.1 --port 8000
```

**Window 2:**
```
cd "Plutus Accountant Automation Suite\frontend"
npm run dev
```

Then go to **http://localhost:5173** in your browser.

---

*For any questions or issues, contact your developer.*
