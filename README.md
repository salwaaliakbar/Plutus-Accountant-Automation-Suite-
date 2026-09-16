# Plutus Accountant Automation Suite — Milestone 1

Automates the monthly bank statement → Excel pipeline for an accountant's practice.

Upload a Barclays bank statement (CSV, PDF or XLSX), and the system:
1. Parses every transaction
2. Assigns UK financial year labels automatically
3. Categorises each transaction using a rule layer + Claude AI
4. Appends all rows to the Excel workbook's RAW tab with live formulas
5. Rebuilds the Analysis pivot with SUMIFS formulas

---

## Prerequisites

| Tool | Version | Download |
|------|---------|----------|
| Python | 3.10 or higher | https://www.python.org/downloads/ |
| Node.js | 18 or higher | https://nodejs.org/ |
| VS Code | Latest | https://code.visualstudio.com/ |

---

## Step 0 — Open a terminal in the project folder

1. Open File Explorer and go into the project folder (the one with `main.py` and `frontend` inside — if you see another folder with almost the same name, open that one first).
2. Click once inside the address bar at the top of the window (where the folder path is written).
3. Type `cmd` and press Enter.

A black window opens — this is your terminal. Do this every time a step below says "open a terminal".

---

## Setup (one-time)

### 1. Install Python packages

Open a terminal (Step 0). Paste this and press Enter:

```
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

### 2. Add your Anthropic API key

Paste this and press Enter:

```
copy .env.example .env
```

Open the new `.env` file, and replace the text with your real key:

```
ANTHROPIC_API_KEY=sk-ant-...your-real-key-here...
```

### 3. Install frontend files

Open a terminal (Step 0). Paste this and press Enter:

```
cd frontend
npm install
```

Did it finish with no red text? Go to the next section.

Did you get a **red error message**? Paste this instead:

```
npm.cmd install
```

---

## Running the project

You need **two terminals open at the same time**.

**Terminal 1 — Backend**

Open a terminal (Step 0). Paste this and press Enter:
```
venv\Scripts\uvicorn api.server:app --host 127.0.0.1 --port 8000 --reload
```

**Terminal 2 — Frontend**

Open another terminal (Step 0). Paste this and press Enter:
```
cd frontend
npm run dev
```

Did you get a **red error message**? Paste this instead:
```
npm.cmd run dev
```

Then open your browser and go to: **http://localhost:5173**

---

## How to use

1. Open **http://localhost:5173**
2. Drag and drop a Barclays bank statement onto the upload zone, or click to browse
3. Wait for the progress steps to complete (Parsing → AI Categorisation → Writing Excel)
4. Click **Download Excel** to get the processed workbook
5. Click **Process another statement** or the back arrow to process another file

---

## Supported input formats

| Format | Source |
|--------|--------|
| `.csv` | Barclays online banking → Download → CSV |
| `.pdf` | Barclays online banking → Download → PDF |
| `.xlsx` | Barclays online banking → Download → Excel |

> The output is always an `.xlsx` file — your master workbook with the new rows appended.

---

## Project structure

```
Plutus Accountant Automation Suite/
│
├── api/
│   └── server.py           # FastAPI backend (upload, progress, download)
│
├── core/
│   ├── financial_year.py   # UK FY logic (6 Apr – 5 Apr boundary)
│   ├── categorise.py       # Rule layer + Claude Haiku categorisation
│   └── excel_writer.py     # Appends to RAW tab, rebuilds Analysis pivot
│
├── parsers/
│   ├── parse_csv.py        # Barclays CSV parser
│   ├── parse_pdf.py        # Barclays PDF parser (pdfplumber)
│   └── parse_xls.py        # Barclays XLSX parser (openpyxl)
│
├── scripts/
│   └── recalc.py           # Optional LibreOffice headless recalculation
│
├── frontend/               # React + Vite + Tailwind frontend
│   └── src/
│       ├── pages/          # UploadPage, ProcessingPage, ResultPage, ErrorPage
│       └── components/     # UploadZone, ProcessingPanel, ResultPanel
│
├── Samples/
│   └── Bank summarised Behesta v1 2025.xlsx   # Master Excel template
│
├── output/                 # Processed files are saved here (auto-created)
├── uploads/                # Temporary upload storage (auto-created)
│
├── main.py                 # CLI entry point (optional, bypasses frontend)
├── requirements.txt        # Python dependencies
└── .env.example            # API key template
```

---

## Column layout written to RAW tab (A–M)

| Col | Name | Content |
|-----|------|---------|
| A | No | Sequential row number |
| B | SA | UK FY label e.g. `FY24` |
| C | ACC | FY range e.g. `24/25` |
| D | Date | Transaction date |
| E | Subcategory | Bank's own transaction type |
| F | Memo | Transaction description |
| G | Paid in | Credit amount |
| H | Paid out | Debit amount |
| I | Balance | Statement balance (blank for CSV/XLSX) |
| J | NET | Formula `=G-H` |
| K | Balance UC | Running balance formula chain |
| L | Check UC | Formula `=I-K` (only when balance present) |
| M | UC Category | Category assigned by the system |

---

## UK Financial Year logic

- Runs **6 April to 5 April**
- Example: transaction on 2 May 2024 → `FY24`, `24/25`
- Example: transaction on 5 April 2024 → `FY23`, `23/24`
- Example: transaction on 6 April 2024 → `FY24`, `24/25`

---

## Category list

The system categorises every transaction into one of these fixed categories:

```
Accountancy, Bank charges, Car Insurance, Charging, Company Car,
DLA, Directors salary, Donation, Entertainment, Equipment,
Gym, HMRC, In/Out, Income, Insurance, Interest income,
Investment, Lunch, Mobile phone, Mother Salary, Parking,
Penalty fee, Petrol, Postage, Professional, Professional fees - College,
Refund, Subscription, Sundry, Taxes for mother, Taxi, Train, Travel
```

Known payees are matched by rules instantly (no API cost). Unknown transactions are sent to Claude AI. Results are cached so repeated payees are never re-sent.

---

## Security notes

- Your Anthropic API key is stored only in the `.env` file — never committed to version control
- Bank account numbers, sort codes, IBANs and account holder names are **never sent to the AI**
- Only the transaction description, subcategory and credit/debit direction are sent for categorisation
- All processing happens on your local machine — no data is sent to any external server except Anthropic's API for categorisation

---

## CLI usage (optional, without the frontend)

You can also run the pipeline directly from the command line:

```
# CSV
venv\Scripts\python main.py --input "Samples\May 24.csv"

# PDF
venv\Scripts\python main.py --input "Samples\Statement 03-MAY-24.pdf"

# Custom template or output path
venv\Scripts\python main.py --input "path\to\statement.csv" --template "path\to\template.xlsx" --output "path\to\output.xlsx"

# Append to RAW (3) instead of RAW (2)
venv\Scripts\python main.py --input "statement.csv" --sheet "RAW (3)"
```

---

## Troubleshooting

**"No transactions found in the uploaded file"**
- Make sure the file is a genuine Barclays export, not a manually created spreadsheet
- Check the file opens correctly in Excel/a text editor

**"No API key – marking X txns as Sundry"**
- Your `.env` file is missing or the key is invalid
- Copy `.env.example` to `.env` and add your key

**Backend won't start**
- Make sure you're in the project root and the venv is activated
- Run `venv\Scripts\pip install -r requirements.txt` again

**Frontend shows blank page**
- Make sure the backend is running on port 8000 before opening the frontend
- Check the browser console for errors (F12)

**`npm install` or `npm run dev` shows a red error about "running scripts is disabled" / execution policy**

Switch to Command Prompt and use `npm.cmd` instead of `npm` — see Setup step 3 and the Running section above for the exact steps.

**Can't find `main.py` / `frontend` / the right folder to run commands in**

See Step 0 above — GitHub downloads sometimes create a folder inside another folder of the same name.
