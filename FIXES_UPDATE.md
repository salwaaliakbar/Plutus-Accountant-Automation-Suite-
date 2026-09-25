# Plutus Bank Statement Processing — Fixes Update

## Summary

All five concerns in your feedback have been fixed and tested against the files you sent. The system now reads bank statements from more banks and formats. It produces the same transactions whether a statement arrives as PDF, CSV or XLSX, and writes complete transaction details into your working papers.

The main changes:

- **PDF statements from more banks.** Starling now works. So do other banks whose PDFs have no table grid lines, such as HSBC, Lloyds, NatWest, Nationwide, Santander and Monzo.
- **Paid in and Paid out always filled.** Amounts land in the right column, whatever the bank or template calls it.
- **CSV and XLSX give identical results.** The same statement now produces the same transactions, in the same order, in either format.
- **Complete transaction details.** Description, type and reference are carried through, so each categorisation can be reviewed.
- **Analysis sheet fixed.** New financial years now show totals, and your own notes on the sheet are no longer deleted.

Every PDF is now checked against the running balance printed on the statement. A clean result proves that no transaction was missed or put in the wrong column.

## Your concerns, one by one

| # | Concern you raised | What was causing it | What we changed | Result |
| --- | --- | --- | --- | --- |
| 1 | Paid in / Paid out missing from PDF output | PDFs without table grid lines could not be read. Templates with column names such as "Money In", "Credit" or "Value" were not recognised, so the amounts were left blank. | A new reader places each amount by where it sits under the column headings. Column names are now matched flexibly ("Paid In (£)", "Money in", "Credit", "Value" and similar). | Amounts land in the correct column for every template tested. |
| 2 | XLSX, CSV and some PDFs (e.g. Starling) giving errors | The system expected column headings on the very first row of the first sheet. Many bank exports put account details above the headings, or a summary sheet first. | The system now finds the heading row wherever it is and skips cover or summary sheets. It also handles semicolon-separated files and more date and amount formats (e.g. "12.00 DR"). | Files with header blocks, cover sheets or different layouts now load correctly. |
| 3 | XLSX and CSV giving different transaction counts | Rows with a date or amount format the system didn't recognise were silently dropped, and which rows that affected differed by format. Newest-first files also came out with each day's transactions in reverse order. | More formats are recognised, and any row that still can't be read is reported instead of silently skipped. Newest-first files are put into date order correctly. | For each client file we tested five variations (XLSX and CSV, with and without header blocks). All 20 gave identical transactions. |
| 4 | Some outputs had amounts only, no other details | Description columns with names such as "Name", "Transaction Details" or "Payee" were not recognised. PDFs also dropped the transaction type. | Many more column names are recognised. If none matches, the most likely description column is used instead of leaving it blank. Transaction type is now carried through from PDFs. | Every tested row has its date, description and amount. |
| 5 | Output inconsistent overall | The combined effect of issues 1 to 4. | All of the above, plus the balance check described in the summary. | The same statement now gives the same, complete result in any format. |

## The files you shared

### Starling statement (PDF, June 2025 to June 2026)

It now reads correctly: all 80 transactions, with type and description. Money in is £60,499.04 and money out is £61,971.57, exactly matching the summary on the statement, and all 63 printed balances reconcile. It had failed because the PDF has no table grid lines, which the old system relied on.

### Suhail Ali (Paid in / Paid out blank)

- **Cause 1:** In this template, Paid in and Paid out are formulas that read the **Value** column. The system didn't recognise "Value" as the amount column, so it left it empty. Paid in, Paid out, NET and the Analysis totals all came out blank as a result.
- **Cause 2:** The bank export lists the newest transactions first, and each day's transactions ended up in reverse order. That made the balance check fail on most rows.
- **Cause 3:** The Analysis sheet mistook the accountant's side notes in its header row for year labels, so its totals were blank.
- **Result:** Tested on the same workbook with all 380 new transactions. Value is filled on every row, Paid in and Paid out calculate correctly, and the balance check is zero on every row. The Analysis sheet shows YE22 to YE26.

### Monzo client (new financial year showed no data)

- **Cause 1:** The template contained about 350 empty pre-numbered rows. The system treated them as data and added the new rows after them. That created the "(blank)" column in the pivot table. It also meant the NET formula was copied from an empty row, so it never reached the new rows.
- **Cause 2:** The Money In and Money Out columns were not recognised. This template also records money out as a negative number.
- **Cause 3:** The Analysis sheet treated "Grand Total", a date and zeros as year labels, and never added the new year (FY26).
- **Result:** New rows now start straight after the real data. All 348 FY26 transactions have Money In or Money Out and the NET formula. The FY26 NET total (£985.41) matches the statement exactly. The Analysis sheet shows FY24, FY25 and FY26, and the pivot table picks up FY26 with no "(blank)" column.

## Further improvements made along the way

While fixing the reported issues we found and fixed these as well:

- **Your notes on the Analysis sheet are kept.** Previously the whole sheet was cleared before the summary table was rebuilt. On one template this deleted the accountant's entire balance-sheet working and their PL/BS tags. Now only the summary table is rebuilt, and everything else stays as it was.
- **Category rows keep their order.** Existing categories stay on the same rows and new ones are added at the bottom, so notes beside each category stay against the right one.
- **New financial years are added safely.** A new year (e.g. FY26) gets its own column in the summary table when that column is free. If you use the next column for your own notes, it is left alone, and the new year appears in the pivot table instead, which updates when the file is opened.
- **Pivot tables are placed in empty space.** The pivot table is moved to columns that are genuinely empty, so it never sits on top of your workings.
- **Your template's sign conventions are followed.** Some templates record money out as a negative number, others as positive. The system now follows whatever your existing rows do. Before, new rows in these templates would have added payments instead of subtracting them.
- **Repeat payments are kept.** Two identical payments on the same day (e.g. two parking charges) are both kept, rather than one being removed as a duplicate.

## What you need to do

- **Reprocess affected clients from the original bank statement.** Upload the statement the client downloaded from their bank, together with the working-papers template. Files produced before this update don't change by themselves.
- **Send us any statement that still doesn't read correctly.** We will run it through the same checks.

## Known limitations

- **Older .xls files** (Excel 97–2003 format) are not accepted yet. Please save them as .xlsx first. We can add direct support if needed.
- **Scanned PDFs** (images with no selectable text) cannot be read reliably. A CSV download from online banking is the best alternative.
- **Warnings** about rows that couldn't be read, or template columns that are missing, are currently recorded on the server, not shown on screen. We recommend showing them in the app as a next step.
