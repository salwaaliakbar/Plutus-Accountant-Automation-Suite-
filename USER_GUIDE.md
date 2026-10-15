# User Guide — using Plutus day to day

This guide is for the person who will actually process bank statements with
Plutus and review its work. It assumes the system is already installed and
running. If it isn't yet, hand this to whoever sets up computers and point
them to **[SETUP_GUIDE.md](SETUP_GUIDE.md)** first — that is a one-time job,
usually done once per computer, not something you repeat for every client.

No programming knowledge is needed for anything in this guide.

---

## Contents

1. [Starting the system](#1-starting-the-system)
2. [Processing a client's statement](#2-processing-a-clients-statement)
3. [Reading the result screen](#3-reading-the-result-screen)
4. [Checking the Excel file it produced](#4-checking-the-excel-file-it-produced)
5. [Correcting a wrong category](#5-correcting-a-wrong-category)
6. [Teaching the system from your corrections](#6-teaching-the-system-from-your-corrections)
7. [Everyday troubleshooting](#7-everyday-troubleshooting)
8. [Quick reference card](#8-quick-reference-card)

---

## 1. Starting the system

Every time you want to use Plutus:

1. Open the project folder (the one with `main.py`, `frontend`, `start.bat` inside).
2. Double-click **`start.bat`**.
3. Two black windows open (the engine and the web page) and, after a few
   seconds, your browser opens automatically at **http://localhost:5173**.
4. Leave both black windows open while you work — closing either one stops
   the system. You can minimise them.

If nothing opens, or you see a red error message, see
[Everyday troubleshooting](#7-everyday-troubleshooting) below, or
[SETUP_GUIDE.md](SETUP_GUIDE.md) if it looks like something was never
installed.

When you are done for the day, close the two black windows (or just leave
your computer on — there's no harm in leaving Plutus running).

---

## 2. Processing a client's statement

You need **two files** for every client you process:

| File | What it is | Where it comes from |
|---|---|---|
| **Bank statement** | The client's transactions for the period you're adding | Downloaded from the client's online banking — CSV, PDF or Excel (`.xlsx` / `.xls`), any UK bank |
| **Template workbook** | That client's existing working-papers Excel file | The file you already have for that client, with their previous months' transactions in it |

Steps:

1. On the **Bank Statement Processor** screen, drag the bank statement onto
   the first box (or click it to browse), and the template workbook onto
   the second box.
2. Click **Process Statement**.
3. Wait while it works through:
   **Uploading → Parsing transactions → AI categorisation → Writing Excel
   file → Complete.**
   This normally takes well under a minute; a very large statement (over a
   thousand transactions) can take a minute or two.
4. When it finishes, you land on the **Done** screen — see the next
   section.

> **Tip:** always use the client's own template — the one that already has
> their earlier months in it, in their own formatting. Plutus reads that
> file to learn the client's categories, their financial-year layout and
> their column names, and appends the new transactions in exactly that
> style. Uploading the wrong client's template will mix up two clients'
> data.

---

## 3. Reading the result screen

Once processing finishes, the screen tells you:

- **Transactions** — how many new rows were added. If this is lower than
  you expected, some rows may already have been in the template (Plutus
  always skips rows it has already added, so re-uploading the same
  statement twice is safe).
- **Categories** — how many different categories were used.
- **A list of categories**, each with a count, so you can sanity-check at a
  glance (e.g. "Lunch 42" looks plausible; "Unknown 42" does not).
- **A yellow "Please check" box** — only appears when there is something
  for you to look at. It can list:
  - Rows where the AI wasn't sure about the category (these are also
    highlighted yellow inside the Excel file itself, see the next
    section).
  - Columns in the template that could not be filled from the statement —
    for example a "Description 2" column, because a PDF statement doesn't
    carry a reference field the way a CSV does.
  - A note that the file was a **scanned PDF**, meaning the figures were
    read from a picture of the page rather than from text. Check these
    numbers especially carefully against the original statement.
  - Balances that don't add up, which usually means a row was misread —
    check that row against the statement.
- A **Download Excel** button — click it to save the finished workbook.
  It downloads with the same name as your statement.

Click the back arrow (**Process another statement**) to do the next
client.

---

## 4. Checking the Excel file it produced

Open the downloaded file in Excel. The new rows are added to the bottom of
the RAW/transactions tab, in the client's own column layout — Plutus never
changes the layout or any of the existing rows.

Two things to look for:

- **Yellow-highlighted category cells.** Hover over one — a note (red
  triangle in the corner) explains: *"Please check: the AI was not sure
  about this category."* This happens when the payee is unfamiliar or more
  than one category is plausible. It is not a sign of a mistake, just a
  flag for a human to confirm.
- **A "Correct category" column**, added automatically next to the AI's
  category column if the template doesn't already have one. This is
  **where you type the right answer** when a category is wrong — see the
  next section.

The Analysis / pivot sheet is also updated automatically: new financial
years get their own column once there are rows to put in them, and your
own notes and workings elsewhere on that sheet are left untouched.

---

## 5. Correcting a wrong category

You do **not** need to edit the category column itself. Instead:

1. Find the row with the wrong category (yellow-highlighted rows are the
   ones most likely to need a correction, but you can correct any row).
2. In the **"Correct category"** column on that same row, type the
   category it should have been.
3. Leave rows that are already correct blank in that column — a blank
   "Correct category" cell means "the AI got it right, nothing to learn
   here."
4. Save the file.

That's it for this file. The correction doesn't do anything by itself yet —
step 6 below is what actually teaches the system.

> You can correct as many or as few rows as you like. Even correcting a
> handful of a client's most common payees (their own salary, a recurring
> supplier, a particular shop) makes a real difference, because the system
> remembers every payee you correct and never needs correcting on that
> payee again for that client.

---

## 6. Teaching the system from your corrections

This is the step that turns your corrections into something Plutus actually
uses on the *next* upload, for *any* client.

1. Find the project's **`Feedbacks`** folder (inside the main project
   folder).
2. Copy your corrected Excel file into that folder (copy, don't move — keep
   your original wherever you normally keep client files).
3. Double-click **`learn.bat`**.
4. A black window opens, reads every file in `Feedbacks`, and prints a
   short report, for example:
   ```
   Feedback files in Feedbacks:
     read     Suhail Ali_processed.xlsx / RAW 25: 8 correction(s)

   Learned 140 payee rules (6 new or changed) -> core\learned_examples.json
     3x  MCDONALDS FROME                       -> Lunch
     2x  EE                                     -> Mobile
     ...
   ```
5. Press any key to close the window. That's the whole process — **you do
   not need to restart the backend**; the very next statement you process
   uses the updated rules.

What counts as a correction:
- A row needs a **description** and something typed in **"Correct category"**
  to be learned from.
- If you type the **same** category the AI already gave (confirming it was
  right), the system still notes this as agreement — it reinforces that
  rule rather than changing it.
- A payee's corrections only become a permanent rule once they agree with
  each other often enough. One contradictory correction among many
  consistent ones won't throw the rule off; a payee you correct to
  different categories each time will be left for the AI to keep
  deciding, since there's no single clear answer to learn.
- You can put as many feedback files in the `Feedbacks` folder as you like,
  and run `learn.bat` again any time — it always re-reads everything in
  that folder from scratch, so there's no harm running it repeatedly.

You can leave your corrected files in the `Feedbacks` folder permanently
(they'll be re-read every time), or clear old ones out once you're
confident the rule has been learned — either is fine.

---

## 7. Everyday troubleshooting

**"All transactions in this statement are already in the workbook"**
You (or a colleague) already processed this exact statement into this
template. Nothing is wrong — there's simply nothing new to add. Check you
picked the right template if you expected new rows.

**A lot of rows are yellow / marked "Unknown"**
Usually means the client has few or no transactions in their history yet
(a brand-new client), so the AI has nothing of theirs to learn from. This
settles down quickly as you correct a few rows and run `learn.bat` — see
sections 5–6.

**"Please check" mentions a column that's empty in the output**
Most likely the bank statement format (often a PDF) simply doesn't contain
that information — for example PDFs don't usually carry a separate
reference number the way a CSV export does. This is expected, not a fault;
fill it in by hand if the template needs it.

**The statement is a photo or scan of a paper statement**
Plutus can still read it, using AI to read the picture rather than the
text. Double-check every figure against the original in this case — the
results page will remind you to.

**The browser window didn't open / shows an error**
Make sure both black windows from `start.bat` are still open and didn't
show a red error when they started. If one of them closed by itself, see
[SETUP_GUIDE.md](SETUP_GUIDE.md)'s troubleshooting section.

**I made a mistake correcting a category — can I undo it?**
Just fix the "Correct category" cell in your file and run `learn.bat`
again before anyone relies on the wrong rule. If a wrong rule has already
been learned, correct it the normal way on the next statement that uses
it, and it will overwrite the old rule once corrections for the right
answer outnumber the wrong one.

---

## 8. Quick reference card

**Every time you process a statement:**
1. `start.bat` (if not already running)
2. Upload statement + template → **Process Statement**
3. Review the result screen, check any "Please check" items
4. **Download Excel**

**When you correct a category:**
1. Open the downloaded file
2. Type the right category in **"Correct category"** on that row
3. Save
4. Copy the file into the **`Feedbacks`** folder
5. Double-click **`learn.bat`**

No restart needed after `learn.bat` — your next upload uses the update
straight away.
