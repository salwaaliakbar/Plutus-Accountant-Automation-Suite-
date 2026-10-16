# Plutus Accountant Automation Suite — Milestone 1

Automates the monthly bank statement → Excel pipeline for an accountant's practice.

Upload a bank statement (CSV, PDF or Excel, from any UK bank) together with
the client's own working-papers template, and the system:
1. Parses every transaction, whatever bank or format it came from
2. Assigns financial year labels in the client's own convention
3. Categorises each transaction using the client's history, learned corrections, and Claude AI
4. Appends all rows to the template's transactions tab with live formulas, matching its existing layout
5. Rebuilds the Analysis / pivot sheet with SUMIFS formulas, without disturbing the accountant's own workings

**New here?**
- **Installing and starting Plutus** → [Getting started](#getting-started) below
- **Teaching it from your corrections** → [Correcting categories and teaching Plutus](#correcting-categories-and-teaching-plutus) below
- **Processing statements day to day, in detail** → **[USER_GUIDE.md](USER_GUIDE.md)**
- From [Supported input formats](#supported-input-formats) onwards, this file is a technical reference (how it works, project layout, CLI usage).

---

## Prerequisites (install once per computer)

| Tool | Version | Download |
|------|---------|----------|
| Python | 3.10 or higher | https://www.python.org/downloads/ |
| Node.js | 18 or higher (the **LTS** version) | https://nodejs.org/ |

> **Installing Python:** on the first screen of the installer, tick **"Add python.exe to PATH"** before clicking **Install Now**. Without it, `install.bat` cannot find Python.

---

## Getting started

Everything is done by double-clicking `.bat` files, so you never need to type commands.

### 1. Download and unzip

1. Download the Plutus `.zip` file.
2. Right-click it, choose **Extract All...**, then click **Extract**.
3. Open the extracted folder. Make sure you are in the folder with `main.py`, `frontend` and the `.bat` files inside. If you only see one folder with almost the same name, open it.

This is your **project folder**. Always use this same folder.

### 2. Install (one-time)

Double-click **`install.bat`**. It sets up everything Plutus needs, which takes a few minutes. Wait until it says **Installation complete**, then press any key to close the window.

> If Windows shows **"Windows protected your PC"**, click **More info** and then **Run anyway**.

### 3. Add your API key (the `.env` file)

The `.env` file holds your Anthropic API key. It must sit in the project folder, next to `main.py`.

- **If you were given a `.env` file:** copy it into the project folder.
- **If not:** `install.bat` creates one and opens it in Notepad. Replace `sk-ant-...` with your real key, then save and close Notepad:

```
ANTHROPIC_API_KEY=sk-ant-...your-real-key-here...
```

> The file name must be exactly `.env`, not `.env.txt`. If the key is missing or wrong, Plutus still runs, but every new payee gets categorised as **Unknown**.

### 4. Start Plutus (every time)

Double-click both of these files in the project folder:

1. **`start_backend.bat`**: opens a window called **Plutus - Engine**.
2. **`start_frontend.bat`**: opens a window called **Plutus - Web UI**, and your browser opens **http://localhost:5173** after a few seconds.

Keep both windows open while you work. To stop Plutus, close both windows.

> **Shortcut:** `start.bat` starts both at once and opens the browser.

---

## How to use

1. Open **http://localhost:5173**
2. Drag and drop the client's bank statement onto the first upload zone, and their existing working-papers template onto the second
3. Click **Process Statement** and wait for the progress steps to complete (Parsing → AI categorisation → Writing Excel)
4. Review the result screen, including any "Please check" items, then click **Download Excel** to get the processed workbook
5. Click **Process another statement** or the back arrow to process another file

For the full walkthrough, see **[USER_GUIDE.md](USER_GUIDE.md)**.

---

## Correcting categories and teaching Plutus

When Plutus gets a category wrong, correct it once and Plutus learns it for next time.

1. Open the processed Excel file you downloaded.
2. Find the wrong category in the **UC category** column. Type the correct category in the column **right next to it**. Leave that cell empty on rows that are already correct.
3. Save the file and put it in the **`Feedbacks`** folder inside the project folder.
4. Double-click **`learn.bat`**. It reads every file in `Feedbacks` and saves what it learned. Press any key to close it when it finishes.

There's no need to restart Plutus. The next statement you upload uses the new corrections.

> **Keep your old files in `Feedbacks`.** `learn.bat` re-learns from every file in the folder each time, so a file you remove takes its corrections with it.
>
> **Be consistent:** Plutus only learns a payee once most of its corrections agree. A payee corrected to different categories each time is left for the AI to decide. See [Learning from corrections](#learning-from-corrections-the-feedback-loop) for the details.

---

## Getting a new version

1. Close the two Plutus windows.
2. Download the new `.zip` and unzip it into a **new** folder, as in step 1 above.
3. From your old project folder, copy these into the new one:
   - the **`.env`** file
   - the files inside the **`Feedbacks`** folder
4. In the new folder, double-click **`install.bat`** (the update may need new packages), then **`learn.bat`** to restore what Plutus learned from your corrections.
5. Start Plutus as usual. Once everything works, you can delete the old folder.

---

## Supported input formats

| Format | Source |
|--------|--------|
| `.csv` | Any UK bank's online banking → Download transactions → CSV |
| `.pdf` | A statement saved as PDF, from online banking or the bank directly (scanned/photographed statements are also read, using AI to read the page image — see Housekeeping below) |
| `.xlsx` / `.xls` | Online banking → Download → Excel (both current and old Excel 97–2003 `.xls` files work) |

> The output is always an `.xlsx` file — the client's own template workbook with the new rows appended, in their existing layout and formatting.

---

## Project structure

```
Plutus Accountant Automation Suite/
│
├── api/
│   ├── server.py           # FastAPI backend (upload, progress, download)
│   └── jobs.py             # Job tracking that survives a restart, file clean-up
│
├── core/
│   ├── financial_year.py   # UK tax-year logic (6 Apr – 5 Apr boundary)
│   ├── year_labels.py      # Learns each client's OWN financial-year convention from their rows
│   ├── columns.py          # Works out which template column plays which role
│   ├── categorise.py       # Client history + learned corrections + Claude AI
│   └── excel_writer.py     # Appends new rows, rebuilds the Analysis / pivot sheet
│
├── parsers/
│   ├── _common.py          # Shared column-detection / row-parsing (CSV & Excel)
│   ├── parse_csv.py        # CSV parser (any UK bank)
│   ├── parse_pdf.py        # PDF parser — ruled tables, column-position layout, AI fallback, scanned PDFs
│   └── parse_xls.py        # Excel parser (.xlsx and legacy .xls)
│
├── scripts/
│   ├── learn_from_feedback.py       # Builds core/learned_examples.json from corrected files in Feedbacks/
│   ├── validate_feedback_accuracy.py # Held-out accuracy check against Feedbacks/
│   └── recalc.py                     # Optional LibreOffice recalculation (not needed for normal use)
│
├── frontend/               # React + Vite + Tailwind frontend
│   └── src/
│       ├── pages/          # UploadPage, ProcessingPage, ResultPage, ErrorPage
│       └── components/     # UploadZone, ProcessingPanel, ResultPanel
│
├── tests/                  # Automated tests (invented data only) — see Housekeeping below
│
├── Feedbacks/               # Put your corrected output files here, then run learn.bat
├── output/                  # Processed files are saved here (auto-created)
├── uploads/                 # Temporary upload storage (auto-created)
│
├── main.py                  # CLI entry point (optional, bypasses frontend)
├── learn.bat                # Reads Feedbacks/ and updates learned_examples.json
├── install.bat              # One-time setup (Python environment, packages, .env)
├── start_backend.bat        # Starts the backend engine (port 8000)
├── start_frontend.bat       # Starts the web interface (port 5173) and opens the browser
├── start.bat                # Starts both at once
├── requirements.txt         # Python dependencies
├── requirements-dev.txt     # Adds pytest and friends, for running the test suite
├── USER_GUIDE.md             # Day-to-day usage guide (start here if you're new to using it)
├── SETUP_GUIDE.md            # First-time installation guide
└── .env.example              # API key template
```

---

## How the template's columns are read

Older versions of this tool expected a fixed column layout. The current
version instead works out each column's role from the client's own
template — by name first, then by the words in the header, then by what the
column actually contains — so it adapts to each accountant's own layout
instead of requiring one fixed format. See
[Working with unfamiliar templates and statements](#working-with-unfamiliar-templates-and-statements)
below for the details, and `core/columns.py` for the implementation.

Typical roles found in a RAW/transactions tab: a sequence number, financial
year label(s), date, description (and sometimes a separate payee column),
reference, transaction type, paid in / paid out (or a single signed
amount), balance, a NET formula, a running-balance check, the category the
system assigns, and a "Correct category" column for the accountant's
corrections.

---

## UK Financial Year logic

- Runs **6 April to 5 April**
- Example: transaction on 2 May 2024 → `FY24`, `24/25`
- Example: transaction on 5 April 2024 → `FY23`, `23/24`
- Example: transaction on 6 April 2024 → `FY24`, `24/25`

---

## Category list

Categories are chosen in this order of priority (see `core/categorise.py`):

1. **The client's own template history** — when a payee has always been
   given one category in that client's existing rows, that category is
   reused directly, with no AI call.
2. **Corrections the accountant made on previous outputs**
   (`core/learned_examples.json`, built by `scripts/learn_from_feedback.py`
   — see [Learning from corrections](#learning-from-corrections-the-feedback-loop)
   below), reused the same way.
3. **The payment reference**, for HMRC payments specifically (PAYE,
   Corporation Tax, Self Assessment, from the reference number's format).
4. **Claude**, for everything else — using the client's own category
   vocabulary where it applies, falling back to a general list when it
   doesn't:

   ```
   AOP, Accommodation, Accountancy, Accountancy fee, Bank charges, Car,
   Car insurance, Car lease, Charging Car, Cleaning, College, Company Car,
   Computer, CPD Grant, DBS, Dinner, Directors Loan Account, DLA,
   Directors salary, Dividends, Donation, Education Course, Entertainment,
   Equipment, FODO, Food, GOC, Gym, Heat, HMRC, HMRC-CT, HMRC-PAYE,
   HMRC-SA, HMRC-VAT, HMRC - maternity pay, Hotel, Income, Insurance,
   Insurance - AOP, Interest income, Internet, Light and heat, Lunch,
   Marketing, Mobile phone, Other direct costs, Parking, PCSE,
   Penalty fee, Petrol, Postage, Professional fee, Professional - College,
   Professional - AOP, Professional - GOC, Purchase, Refund, Rent,
   Repairs, Sales, Security, SIPP, SMP, Staff benefits, Stationary,
   Subscriptions, Sundry, Taxi, Telephone, Toll, Train, Transfer, Travel,
   Trivial benefits, Wages, Wages-staff, Water, Website, Work from home,
   Unknown
   ```

   Claude also says whether it is sure; rows it isn't sure about (and any
   row it can only answer "Unknown" for) are highlighted yellow in the
   output for the accountant to check.

Afterwards, every category is reduced to one spelling per meaning (`DLA`
vs `Directors Loan Account`) — preferring the client's own spelling — and
payments to the same person for the same purpose (e.g. several wordings of
a director's loan) are given one consistent category.

Known payees are matched instantly, at no API cost. Only genuinely new
payees are sent to Claude, and identical transactions within one statement
are only asked about once.

---

## Learning from corrections (the feedback loop)

When the accountant types the right category into the **"Correct
category"** column of a processed file and puts that file in the
`Feedbacks/` folder, running `learn.bat` (or
`venv\Scripts\python scripts\learn_from_feedback.py`) rebuilds
`core/learned_examples.json` from every file in that folder. No restart is
needed — the backend reads the file fresh on each upload.

A payee's corrections become a rule once they agree with each other at
least 60% of the time (`MIN_SHARE` in `scripts/learn_from_feedback.py`); a
payee corrected to different categories each time is left for the AI to
keep deciding, since there's no single clear answer to learn from it.

See **[USER_GUIDE.md](USER_GUIDE.md#5-correcting-a-wrong-category)** for
the full step-by-step walkthrough, and
`scripts/validate_feedback_accuracy.py` for a held-out accuracy check
against the `Feedbacks/` files.

---

## Housekeeping and tests

- Processed workbooks, job records and uploads are deleted after 24 hours (change with the `PLUTUS_RETENTION_HOURS` environment variable). Finished jobs can still be downloaded after the server restarts.
- Scanned (image-only) PDFs are read by Claude from the page images, every row is checked against the printed running balance, and the results page tells the accountant to verify the figures. **Privacy:** for scans only, the page images (including anything printed on them, such as name and account number) are sent to Anthropic. Set `PLUTUS_SCAN_AI=0` in `.env` to switch this off, in which case scans are rejected with a clear message.
- Excel recalculates the output when it opens it, so LibreOffice is not needed. For the CLI, `--recalc` is optional and never replaces the file if it would lose PivotTables.
- Run the tests (all use invented data): `venv\Scripts\pip install -r requirements-dev.txt` then `venv\Scripts\python -m pytest`.
- The Claude model used for categorisation can be changed with the `CATEGORISE_MODEL` environment variable (default `claude-sonnet-5`); see `.env.example` for this and the other optional settings above.

---

## Working with unfamiliar templates and statements

- **Template columns are recognised by name, then by the words in the header, then by what the column contains** (see `core/columns.py`). So `UC category`, `Category  UC` and `UC desc` are all found as the category column; `Year` / `Personal` holding `FY24` / `23/24` are found as year columns; `Spending`, `Category Bank` and a plain `Category` next to a UC column are the bank's own category.
- **Two-column review templates** (`UC desc` = first pass, `MUC desc` = the reviewer's final category): the AI fills the first-pass column only and leaves the reviewer's column empty, learns from the reviewer's column, and builds the Analysis on whichever category column the Analysis labels actually use.
- **Any statement column whose header matches a template header is copied across** (Transaction ID, Time, Category, Account Name ...).
- **Blank Type / Name / bank-category cells** (PDF statements do not print them) are filled from what the client's own rows hold for the same payee, in the client's own spelling — only when that payee repeated the same value at least twice. Nothing is guessed.
- **Anything that still cannot be filled is reported** on the results page ("Please check") instead of being left blank silently.
- **Spreadsheets saved with day and month swapped** (US-locale Excel) are detected from the order of the dates and put right.
- **HMRC payments** are categorised from the payment reference format (PAYE, Corporation Tax, Self Assessment) — unless the client's own history already has a consistent category for that payee, which always wins.
- **A separate payee column** (Monzo's `Name` next to a `Description` that is often just a code) is used for matching history and sent to the AI, so `EE` / `BP` / `O2` are recognised.
- **"Food?"-style labels** in the accountant's own rows are treated as *Food, please check*: the category written is `Food` and the row is highlighted for review. The question-mark spellings are never offered to the AI as categories.
- **The Analysis grid is only rebuilt on the RAW year column whose labels match its headings**; when no column matches, the sheet is left untouched.

---

## Security notes

- Your Anthropic API key is stored only in the `.env` file — never committed to version control.
- **Account numbers, sort codes and IBANs are scrubbed from every field before it is sent for categorisation** (see `_scrub` in `core/categorise.py`). This is a pattern-based scrub for those specific number formats — it does not remove ordinary names that appear in a transaction description or payee field (e.g. a payment reference naming the person paid), since that wording is usually what lets the AI categorise the payment correctly (salary vs dividend vs loan, for example).
- The fields sent for categorisation are the date, description, payee (when the statement has one), transaction type, payment reference, the bank's own category hint, and the money in/out amount — all scrubbed as above first.
- All processing happens on your local machine — no data is sent to any external server except Anthropic's API, for categorisation and, for scanned PDFs only, reading the page images (see [Housekeeping and tests](#housekeeping-and-tests)).

---

## CLI usage (optional, without the frontend)

You can also run the pipeline directly from the command line. This project
does not ship a sample template, so always pass `--template` explicitly,
pointing at the client's own working-papers file:

```
# CSV
venv\Scripts\python main.py --input "path\to\statement.csv" --template "path\to\template.xlsx"

# PDF (including a scanned PDF, if ANTHROPIC_API_KEY is set)
venv\Scripts\python main.py --input "path\to\statement.pdf" --template "path\to\template.xlsx"

# Custom output path
venv\Scripts\python main.py --input "path\to\statement.csv" --template "path\to\template.xlsx" --output "path\to\output.xlsx"

# Append to a specific RAW sheet when the template has more than one
venv\Scripts\python main.py --input "statement.csv" --template "template.xlsx" --sheet "RAW (3)"

# Also recalculate with LibreOffice (optional; Excel recalculates on open anyway)
venv\Scripts\python main.py --input "statement.csv" --template "template.xlsx" --recalc
```

---

## Troubleshooting

> For everyday issues while processing statements (a wrong category, a
> "Please check" message, a statement that's already fully in the
> workbook, a scanned PDF), see
> **[USER_GUIDE.md § 7](USER_GUIDE.md#7-everyday-troubleshooting)**. The
> items below are about installing or running the system itself.

**"No transactions found in the uploaded file"**
- Make sure the file is a genuine export from online banking, not a manually created spreadsheet
- Check the file opens correctly in Excel/a text editor

**"No API key — marking N txns as 'Unknown'"**
- Your `.env` file is missing or the key is invalid
- See [step 3 of Getting started](#3-add-your-api-key-the-env-file): put the `.env` file next to `main.py` with your real key, then close and restart `start_backend.bat`

**A `.bat` window says "Python environment not found", "Web interface not installed" or ".env file not found"**
- Setup hasn't been done in this folder yet. Double-click `install.bat`, wait for **Installation complete**, then try again.
- For `.env file not found`, see [step 3 of Getting started](#3-add-your-api-key-the-env-file).

**`install.bat` says "Python not found" or "Node.js not found"**
- Install the missing tool (see [Prerequisites](#prerequisites-install-once-per-computer)). For Python, tick **"Add python.exe to PATH"**. Then double-click `install.bat` again.

**Backend won't start**
- Double-click `install.bat` again. It is safe to re-run and reinstalls any missing packages.

**Frontend shows blank page**
- Make sure the backend is running on port 8000 before opening the frontend
- Check the browser console for errors (F12)

**Can't find `main.py` / `frontend` / the `.bat` files**

Unzipping sometimes creates a folder inside another folder of the same name. Open the inner folder. See [step 1 of Getting started](#1-download-and-unzip).
