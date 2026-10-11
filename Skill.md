---
name: importquote
description: Extract the quote number, date and line items (Part ID, description, U/M, unit price, qty) from an ACE Engineering quote PDF and write them into the standard Quote layout (GetQuoteParams + GetQuoteData tables) on an Excel sheet.
---

## When to use
- The user uploads one or more supplier quote PDFs (ACE Engineering layout: "Quote no NNNNN", "Line N Part ID …" blocks) and asks to import, extract or load the quote into Excel.
- The user runs /importquote, optionally with a target sheet name.
- Not for quotes in a different layout. If the PDF has no "Line N Part ID" blocks, stop and tell the user the layout is not recognised.

## Source PDF layout (what to expect)
- Every page header: company name + 3 address lines, then `Quote no <number>` and the date on its own line (d/mm/yyyy, NZ format). Every page footer: `N of M`. These repeat on every page and must **not** leak into descriptions.
- Each item block:
  - `Line <n> Part ID <PartID>`
  - 1+ description lines (a block can break across a page: continue collecting after the page header)
  - `Unit price Discounted Totals with GST` / `Qty U/M excl GST unit price Excluded` (column headings, skip)
  - Values line: `<Qty> <U/M> <Unit price excl GST> <Discounted unit price> Extended <extended>`
  - `Discounted <amount>` (may land on the next page, ignore)
- End: `Quote totals excl GST GST Total`, then `Full price …` and `Discounted price …`.
- Known text-extraction glitch: digits sometimes have stray spaces ("2 5.00" = 25.00, "1 7.50" = 17.50). Remove spaces before parsing numbers. Every amount has exactly 2 decimals.

## Step 1: Parse the PDF (code_execution, Python)
Find uploads with `glob.glob(os.environ.get('INPUT_DIR','/files/input') + '/**/*.pdf', recursive=True)` and parse each with this tested parser:

```python
import re, pdfplumber
NUM = r'\d[\d,]*?\.\d{2}'
def nums(s): return [float(x.replace(',', '')) for x in re.findall(NUM, s.replace(' ', ''))]
def parse_quote(path):
    with pdfplumber.open(path) as p:
        lines = [l.strip() for pg in p.pages for l in (pg.extract_text() or '').splitlines()]
    qno = re.search(r'Quote no\s+(\d+)', ' '.join(lines)).group(1)
    qdate = next(l for l in lines if re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', l))
    skip = set(lines[:5]) | {qdate}          # page-header lines + date
    items, cur, state, full = [], None, None, None
    for l in lines:
        if l in skip or re.fullmatch(r'\d+ of \d+', l) or l.startswith('Qty U/M'):
            continue
        m = re.match(r'Line\s+\d+\s+Part ID\s+(\S+)', l)
        if m:
            cur = {'PartID': m.group(1), 'desc': []}; items.append(cur); state = 'desc'; continue
        if l.startswith('Unit price Discounted'): state = 'vals'; continue
        if l.startswith('Quote totals'): state = 'totals'; continue
        if state == 'desc':
            cur['desc'].append(l)
        elif state == 'vals':
            m = re.match(r'^(\d+)\s+([A-Za-z]+)\s+(.*?)\s+Extended\s+(.*)$', l)
            if m:
                u = nums(m.group(3))
                cur.update(Qty=float(m.group(1)), UM=m.group(2), Unit=u[0], Disc=u[1], Ext=nums(m.group(4))[0])
        elif state == 'totals' and l.startswith('Full price'):
            full = nums(l)[0]
    for i in items: i['Description'] = ' '.join(i.pop('desc'))
    return qno, qdate, items, full
```

Checks (print only these, never the full text):
- Every item has Qty, UM, Unit. Print item count.
- For each item, `Qty*Unit == Ext` (±0.01).
- `sum(Qty*Unit) == Full price` excl GST (±0.01). If any check fails, report the line and do not write until resolved.
Done when: quote no, date, and all items parsed and both checks pass.

## Step 2: Choose the destination
- Default target: `Sheet1`. If the user named a sheet, use it.
- If the target already contains a populated `GetQuoteData` table for a **different** quote, tell the user which quote is there and ask: replace it, or put the new quote on a new sheet named `Quote <number>`. Don't overwrite until they confirm (unless they said "replace"/"overwrite").
- Table names must be unique per workbook. On Sheet1 use `GetQuoteParams` / `GetQuoteData`; on a new sheet use `GetQuoteParams_<no>` / `GetQuoteData_<no>` and adjust formula references to match.
- If several PDFs are uploaded, import each to its own sheet `Quote <number>`.

## Step 3: Build or refresh the layout (execute_office_js)
Exact layout from the user's examples (font Aptos Narrow 12 throughout, TableStyleMedium7 for both tables):

| Cell / range | Content |
|---|---|
| A2:C3 | Table `GetQuoteParams`, headers `Quote no`, `Quote date`, `PDF filename`. A3 = quote no **as text** (numberFormat `@`), B3 = date as a real date serial (convert d/mm/yyyy, format `dd-mm-yy`, DMY order), C3 = PDF file name |
| A11 | `Quote` bold 16pt |
| A12 / B12 | `Quote no` (bold) / `=GetQuoteParams[Quote no]` left-aligned |
| A13 / B13 | `Quote date` (bold) / `=GetQuoteParams[Quote date]` format `dd-mm-yy` (DMY), left-aligned |
| F13 | `Prices excl GST`, right-aligned |
| A14 / B14 | `Quote file` (bold) / file name, font colour #0070C0, left-aligned. Workbook name `deQuoteFilename` → `=Sheet1!$B$14` (add only if missing) |
| A15:F15 | Table `GetQuoteData` headers: `PartID`, `Description`, `U/M`, `Unit price`, `Qty`, `Total price`; A15 bold |
| Data rows | A PartID (text), B full description, C U/M, D Unit price (**undiscounted** unit price excl GST), E Qty |
| F (calc column) | `=[@[Unit price]]*[@Qty]` |
| Totals row | ON. A = `Total`, F = `=SUBTOTAL(109,GetQuoteData[Total price])` |

Formats: D and F `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`; E `_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)`. Column B wrap text, rows autofit. Column widths (chars): A 12.1, B 40.6, C 7.3, D 11.6, E 5.8, F 11.6. Freeze panes: `freezeRows(15)` and `freezeColumns(2)`. Set print area to the GetQuoteData table if building new.

Refreshing an existing layout (tables already exist):
1. Write A3:C3 of GetQuoteParams and B14.
2. Resize GetQuoteData to header + N data rows (`table.resize("A15:F" + (15+N))`; totals row moves automatically), clear old body contents in A:E, write the N rows to A:E, then re-apply the F formula to the whole data body column so it is consistent.
3. Clear any leftover cells below the totals row that belonged to the previous, longer quote.

Building from scratch (no tables): write the labels/headers, create both tables with `sheet.tables.add(range, true)`, set names and styles, write data, enable `showTotals`, then apply formats above. Suspend calculation while writing and restore it in a `finally` block.

Provenance: put one note on the GetQuoteData top-left header cell (A15): `Source: <file name>, uploaded <yyyy-mm-dd>`. No notes on value cells.
Done when: tables exist with correct names and N data rows.

## Step 4: Verify
Read back A2:F(16+N) with get_cell_ranges and confirm:
- B12/B13 show the quote no/date, and the date is the right day (d/mm source → no day/month swap: 1/04/2026 = 1 April 2026).
- Row count = item count; no description contains the quote date, address lines or "N of M".
- Totals row F equals the PDF "Full price excl GST" figure.
- No #VALUE!/#REF!/#NAME?.

## Step 5: Report
One short paragraph: quote no, date, number of lines, total excl GST (and that it reconciles to the PDF), where it was written. Mention anything odd (e.g. a line with no values, discounted prices not imported).

## Notes and design choices
- Descriptions are the full text of all description lines joined with single spaces. The user's original example outputs were missing some description lines and included page-break remnants (e.g. "1/04/2026" inside a description); this skill deliberately keeps every description line and strips page header/footer text.
- Discounted unit price is parsed (`Disc`) but not written, matching the examples. If the user asks, add a `Discounted price` column after `Unit price`.
- The PDF filename is the uploaded file's name (the add-in can't see the user's original local folder path); if the user gives the full path, use that instead.

## Example
Input `Quote1.pdf` → A3 `44556`, B3 1 Apr 2026, 5 rows: ER10008255K 43,750.00; ER10001105 550.00; ER10001798K 1,945.00; E50001509 521.00; ER10000118K 25.00 → Total $46,791.00 (= PDF Full price excl GST).
Input `Quote2.pdf` → `44720`, 22 Apr 2026, 5 rows, Total $4,147.50.