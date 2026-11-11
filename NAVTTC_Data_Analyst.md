# NAVTTC Data Analyst: Excel Handbook and Skill Guide

*A practical Excel data-analyst playbook, written from the real tasks, files and prompts used in the sessions, plus the
general Excel expertise behind them.*

**What this file is.** The full, human-readable version of the `navttc-data-analyst` skill. The skill file
(`NAVTTC_Data_Analyst.skill`) holds the same content split into an operating guide, six reference documents and two scripts, so an AI assistant loads
only what each task needs. This markdown file puts everything in one place for people.

**Skill description (what makes it trigger):**
> Practical Excel data-analyst assistant for NAVTTC-style work. Use this skill whenever the user shares an .xlsx or .csv file or describes a spreadsheet and wants formulas, a lookup, a validation rule, duplicate or blank-cell analysis, a pivot table, a chart, a one-page dashboard, or plain-English sales insights. Trigger on requests such as "write an Excel formula", "update the sheet", "apply this formula", "when I type a code show the whole row", "validate the email column", "count duplicates and null cells", "make a pivot table", "make the graph better", "design a sales dashboard", "fetch rows where sales are above X", or "write 5 insights for a sales manager", even if the user never says the word "skill" or "dashboard". It covers the full loop of inspecting the file, building with live formulas, testing against an independent calculation, and delivering the updated file with a short plain-English summary.

## What is inside

| Part | Content |
|---|---|
| 1 | Operating guide: how to work on any Excel task, rules and reasons, playbooks, answer shapes |
| 2 | Worked examples: the 12 tasks built in the sessions, with formulas, results and lessons |
| 3 | Prompt templates: the six-part prompt, a prompt library, the original prompts cleaned up, the manager-insights prompt |
| 4 | Formula cookbook: tested formulas, validation and conditional-format rules, compatibility |
| 5 | Dashboards and pivot tables: design, building real pivots by code, the calc-sheet pattern, colours |
| 6 | Data quality, pitfalls and the verification protocol |
| 7 | Excel expertise notes: tool choice, functions, errors, speed, auditing, cleaning, statistics, dates, shortcuts, Power Query |

**Reading tip.** Part 1 was written for the skill, so it points to files such as `references/formula-cookbook.md`. In this handbook those are: worked-examples = Part 2, prompt-templates = Part 3, formula-cookbook = Part 4, dashboards-and-pivots = Part 5, data-quality-and-pitfalls = Part 6, excel-expertise = Part 7.

## Files produced in the sessions

| File | What it holds |
|---|---|
| `NIB_Analysis.xlsx` | Bank-records analysis: branch counts, duplicate names, address and CNIC uniqueness, blank cells, 17 charts |
| `Book_4_Sheet1.xlsx` / `Book_4_Sheet1_updated.xlsx` | Bonus formula, North-March total, sales above a threshold, highlight of the top sale, sorted chart data, dashboard |
| `products_updated.xlsx` | Type a product code in Sheet2 and the full row appears from Sheet1 |
| `prectice_record_updated.xlsx` | CNIC code before the dash fills the region from earlier records |
| `prectice_record_updated_email.xlsx` | Email must end with @buitms.edu.pk; validation, check column, colours |
| `sales_with_pivot.xlsx`, `piovt_table_sheet_with_pivot.xlsx` | Real PivotTables: monthly sales by region |
| `Shoes_Business_Tracker_with_Dashboard.xlsx` | One-page sales dashboard with working filters, plus repairs to broken links and `MAXIFS` |
| `NAVTTC_Data_Analyst.skill` | The skill package (SKILL.md, 6 references, 2 scripts) |

## Scripts in the skill package

- `scripts/profile_data.py`: profile a workbook or CSV (empty cells, placeholder text, duplicates, types, footer rows).
  `python scripts/profile_data.py file.xlsx --header-row 2`
- `scripts/make_pivot.py`: build a real PivotTable through LibreOffice.
  `python scripts/make_pivot.py --src in.xlsx --out out.xlsx --rows date --cols region --pages product --data sales:sum --group date`

Both were tested on the sessions' own files. For example, `profile_data.py` reproduced the figures found by hand (337 placeholder addresses, 107 duplicate names covering 291 records,
60 branches), and `make_pivot.py` rebuilt the monthly-sales-by-region pivot with the correct grand total of 8,156,003.

---
# Part 1. Operating guide

Act as a careful, practical Excel analyst. The user is usually a learner or office worker with a real file, a short message
(often with typos), and a need for something that works the first time. Be direct, deliver the file, explain simply, and prove
that it works.

This skill packages the method, the tested formulas, and the lessons from a long run of real tasks: duplicate and blank-cell analysis, lookups,
region-from-ID formulas, email validation, pivot tables, and a one-page sales dashboard.

## How to work (every task)

1. **Read the file first.** Look at the sheets, the header row, the data size, footer or total rows, existing formulas, validation and
   formatting. Run `scripts/profile_data.py <file> --header-row N` for a quick profile. Do this before asking anything, because the
   file usually answers the question.
2. **Say how you read the request, in one line.** Typos and shorthand are normal ("rigon", "cince", "fourmula"). Infer the meaning,
   state it ("taking 'identity' to mean the CNIC column"), and carry on. Ask a question only when two readings would give very
   different results and the file cannot decide. When the user says "assume X", just assume X and say so.
3. **Pick the approach that works everywhere.** Use `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `IFERROR`, `SUMPRODUCT`, `TEXT`. Avoid
   `FILTER`, `UNIQUE`, `XLOOKUP`, `LET` unless the user is on Excel 365. If the user asked only for a formula, give the formula;
   if they asked for "the updated file", deliver the file.
4. **Build with live formulas.** Never paste calculated results where a formula can do the job. Put thresholds and names in labelled cells.
   Add helper columns or a helper sheet rather than writing one giant formula. Keep the user's data and layout; add new columns
   or sheets instead of overwriting.
5. **Test it for real.** Recalculate and make sure there are zero errors, then compare the key numbers against an independent
   calculation, then try edge cases (blank, unknown code, no data, a new row). A clean recalculation alone does not prove the numbers are right.
   See `references/data-quality-and-pitfalls.md`, section 5.
6. **Look at it.** Convert new sheets to an image and check layout, labels and charts.
7. **Deliver and summarise.** Give the file, then a short summary: what was built, the headline numbers, any assumptions, any problems found in the data,
   and anything that could not be tested.

## Rules and the reasons behind them

- **Live formulas, not hard-coded values.** Data changes; hard-coded answers go stale silently.
- **Do not change the user's data.** Prove it with a before/after comparison of every original cell. Do not add example rows to their sheets; test on a copy.
- **Blank, placeholder and zero are different.** "NA" or "ADDRESS MISSING" is text that means "missing". Count it separately from empty cells.
  Read files with `openpyxl` (not default pandas, which turns "NA" into blank).
- **Tell the truth about what you could not do.** Example: PivotTables built through LibreOffice were not opened in desktop Excel; say so.
- **Format answers to the request.** "One formula + 2-line explanation" means exactly that. Lead with the answer; caveats after.
- **Use exact numbers and plain words.** "Rs 173,989", not "about 174k"; "fell", not "contracted"; define "unique", "distinct", "duplicate", "null" when used.
- **A CSV cannot hold formulas.** Deliver .xlsx and say why.
- **Match the workbook.** Reuse its fonts, header colours and sheet names; blue text for cells the user types in; a short "How to use" note where others will fill the sheet.
- **Fix what you find, and say so.** If you hit a real fault in the user's file (a broken link, a `#NAME?` from a missing `_xlfn.` prefix, a dead drop-down),
  fix it minimally and list the fix. Do not rewrite unrelated parts.

## Task playbooks

| The user wants | Do this | Read |
|---|---|---|
| A formula | Give the formula in a code block, then the explanation in the requested length. State assumptions (range, year, name). | `references/formula-cookbook.md` |
| "Update the sheet" | Add the formula or column to a copy of their file, recalculate, check, and return the file. | cookbook |
| Type a code, see the whole row | `INDEX/MATCH` with whole-column ranges, a status column ("OK", "Not found", "Duplicate"), a drop-down, formulas filled well past the data. | cookbook section 3 |
| Fill a field from earlier records | Lookup range that ends one row above the formula (never circular); a visible flag for new values. | cookbook section 3; worked example 8 |
| Validate input (email, ID, range) | Data validation (custom formula, stop alert, input message with an example) **plus** a check column and conditional formatting. | cookbook section 8 |
| Rows above/below a limit | Helper match-number column plus `INDEX/MATCH`; limit in a labelled cell. | cookbook section 4 |
| Highlight the highest or lowest | Conditional formatting with a formula; never colour by hand. | cookbook section 9 |
| Duplicates, unique values, blank cells | Clean keys, counts, first-occurrence flag, status; define the words; report placeholders separately. | cookbook section 6; data-quality |
| Pivot table | Say what goes in Rows, Columns, Values, Filters. If asked to build it, use `scripts/make_pivot.py`. | `references/dashboards-and-pivots.md` |
| Charts | Native Excel charts; combo charts for trend plus profit; a second view when one value dominates. | dashboards |
| One-page dashboard | KPI cards, drop-down filters, a calc sheet of flags, four charts, a data-check line, one printed page. Test several filter combinations. | dashboards sections 3-7 |
| Insights for a manager | 5 bullets, exact numbers, the biggest change named, one action, no jargon. If the table is missing, ask for it. | `references/prompt-templates.md` section 5 |
| Help writing a good prompt | Offer the six-part template (Role, Context, Layout, Task, Format, Example). | prompt-templates |
| Something is broken | Find the cause (error table, type problems, text dates, wildcards), fix, retest. | `references/excel-expertise.md` section 3 |

## Tools in this skill

- `scripts/profile_data.py` : profile a workbook or CSV (empty cells, placeholder text, duplicates, types, footer rows).
  Example: `python scripts/profile_data.py file.xlsx --header-row 2`
- `scripts/make_pivot.py` : create a real PivotTable (needs LibreOffice).
  Example: `python scripts/make_pivot.py --src in.xlsx --out out.xlsx --rows date --cols region --pages product --data sales:sum --group date`

For recalculation use LibreOffice headless (in the Claude sandbox, `recalc.py` from the xlsx skill if present). For writing workbooks use `openpyxl`.
openpyxl cannot create PivotTables or slicers, and it drops existing charts and drawings when it re-saves a file, so keep that in mind when editing a
file that already has charts.

## Reference index (read only what the task needs)

- `references/worked-examples.md` : 12 finished tasks with formulas, results and lessons. Read when the new request resembles one.
- `references/formula-cookbook.md` : tested formulas by task, validation and conditional-format rules, compatibility notes.
- `references/dashboards-and-pivots.md` : pivot design, building pivots by code, dashboard layout, colours, the calc-sheet pattern.
- `references/data-quality-and-pitfalls.md` : profiling, blanks versus placeholders, a table of pitfalls, Google Sheets quirks, the verification protocol.
- `references/excel-expertise.md` : general Excel knowledge: tool choice, function families, errors, speed, auditing, cleaning, statistics, dates, shortcuts, Power Query.
- `references/prompt-templates.md` : the six-part prompt, a prompt library, the original prompts cleaned up, and the manager-insights prompt with an example.

## Answer shapes

**Formula only**
```
=<formula>
```
Two lines: what it does; what it assumes.

**Delivered file**
1. One sentence on what was built and where (sheet and cell).
2. The headline numbers.
3. What to know: assumptions, data problems found, repairs made.
4. One line on what was not tested, if anything.
Keep it short; the file is the deliverable.

**Insights**
Five bullets, one with the biggest change (both values, the gap, the percent), then "Action:" and one concrete step.

---

# Part 2. Worked examples: everything built in the practice sessions

Each entry lists the ask, what was delivered, the key technique, the real result, and the lesson. Use these as
patterns for new requests.

## Contents
1. NIB bank records: counts, duplicates, blanks, charts
2. Bonus formula
3. North sales in March (six-part prompt)
4. Sales above a threshold, pulled into a new sheet
5. Highlight the highest sale and chart all sales
6. A dashboard that is "not a simple chart"
7. Product code lookup across two sheets
8. CNIC code to region
9. Email validation
10. Pivot table: monthly sales by region (advice, then built)
11. Shoes Business one-page sales dashboard
12. Insights for a sales manager

---

## 1. NIB bank records: counts, duplicates, blanks, charts
**Ask:** Count records per branch; count duplicate names; unique vs duplicate addresses and identity numbers; count blank cells
in the whole sheet; full analysis with graphs.
**Delivered:** `NIB_Analysis.xlsx` with Summary, Branch_Count, Name_Duplicates, Address_Duplicates, CNIC_Duplicates,
Null_Analysis, a helper sheet `Row_Flags` (one row per record), and the untouched original `NIB` sheet. 41,361 formulas, 17 native charts.
**Key technique:**
- Normalised keys (`UPPER(TRIM())`), a count per key, a first-occurrence flag, and a status of Unique / Duplicate / Blank / Placeholder.
- Metrics come from `COUNTIF`/`COUNTIFS` on the status columns, so every number is a live formula.
**Result:** 2,512 records, 60 branches (Gul Tower Branch Karachi largest: 300, 11.9%). 107 duplicate names covering 291 records.
CNIC empty in 1,082 records; 33 CNICs repeat. Addresses: 2,045 unique, 50 repeat (125 records), 337 placeholders
("NA" or "ADDRESS MISSING"), 5 truly empty. 25,212 empty cells in the data table (41.8%); 28,181 in the whole used sheet.
**Lessons:**
- Read the file with openpyxl, not pandas. pandas silently turned the text "NA" into blanks, which hid 322 placeholder addresses.
- A blank cell and the text "NA" are different things. Report both.
- The sheet had 16 footer total rows. Exclude them from record counts and say so.
- A `nan` value written into a text key became the "name" `NAN` and created a false duplicate. Test counts against an independent calculation.
- "Identity" was interpreted as the CNIC column, and that interpretation was stated.

## 2. Bonus formula
**Ask:** Formula that returns Bonus if sales are 50,000, else No bonus; then the updated sheet.
**Formula:** `=IF(A2>=50000,"Bonus","No bonus")`, filled down.
**Delivered:** `Book_4_Sheet1.xlsx` (a CSV cannot keep formulas, so the file was saved as .xlsx).
**Lesson:** "Sales are 50000" is ambiguous. Use `>=`, say so, and give the `=` version in one line.
All 8 values were below 50,000, so every row said "No bonus". Say that out loud instead of leaving the user to wonder.

## 3. North sales in March (six-part prompt)
**Ask:** Role / Context / Layout / Task / Format / Example. One formula plus 2-line explanation.
**Formula:** `=SUMIFS(D2:D2001, B2:B2001, "North", A2:A2001, ">="&DATE(2026,3,1), A2:A2001, "<"&DATE(2026,4,1))`
**Delivered:** `Book_4_Sheet1_updated.xlsx` with the formula in G2. It returns 0 on the sample because the data had no North rows and only January dates.
**Lessons:**
- Use a date range (`>=` first day, `<` first day of next month) instead of `MONTH()`. It is faster and needs no array formula.
- CSV dates such as 1/13/2026 become text in day-first regional settings. Mention "real Excel dates".
- Honour the Format line: formula plus exactly two lines.

## 4. Sales above a threshold, pulled into a new sheet
**Ask:** Fetch every row with sales greater than 50,000 (date, region, product, sales) and update the sheet.
**Delivered:** A `High_Sales` sheet driven by a threshold cell (B3). Only one row matched: 1/14/2026, Multan, atta, 544,544.
**Technique (no FILTER function):**
- Helper column on the data sheet numbers the matching rows: `=IF(AND(ISNUMBER(D2),D2>High_Sales!$B$3),COUNTIF(D$2:D2,">"&High_Sales!$B$3),"")`
- The result table pulls row k with `INDEX/MATCH` on that number: `=IF($A7="","",INDEX(Sheet1!A$2:A$2001,MATCH($A7,Sheet1!$E$2:$E$2001,0)))`
- Count of matches: `=COUNTIF(Sheet1!$D$2:$D$2001,">"&$B$3)`; the table has 500 slots and a warning if it overflows.
**Lessons:** `FILTER` needs Excel 365 and does not recalculate in many tools. A helper number plus `INDEX/MATCH` works everywhere.
Test with a lower threshold (1,000 returned the expected 5 rows).

## 5. Highlight the highest sale and chart all sales
**Delivered:** Orange highlight on the top row using conditional formatting: `AND(ISNUMBER($D2),$D2=MAX($D$2:$D$2001))`.
A `Sales_Chart` sheet with sales sorted highest to lowest using `RANK` with a tie-break:
`=RANK(I6,$I$6:$I$17)+COUNTIF($I$6:I6,I6)-1`, then `INDEX/MATCH` on that rank.
**Lesson:** One sale (544,544) was 97.0% of the total, so a normal column chart hid everything else. Show a second view
(log scale, or "everything except the top sale").

## 6. A dashboard that is "not a simple chart"
**Delivered:** A `Dashboard` sheet with 8 KPI tiles (total 561,211; highest 544,544; average 46,768; median 555 and so on), a Pareto chart (columns plus
cumulative-% line on a second axis), a daily-sales-plus-running-total combo with the highest day in orange, an "everything except the top sale" bar chart, and
formula-driven insight sentences.
**Lessons:**
- Highlight the maximum dynamically with a second series (`=IF(F49=MAX(F$49:F$60),F49,0)`) overlapped on the first, so the colour moves when the data changes.
- Median versus average tells the manager how skewed the data is (555 vs 46,768).
- Insight sentences built with `TEXT()` and `IF()` stay correct when data changes.

## 7. Product code lookup across two sheets
**Ask:** Type a product code in Sheet 2 and see the complete row from Sheet 1, for all products.
**Delivered:** `products_updated.xlsx`. Codes in Sheet2!A; detail columns B:F by formula; a status column.
**Formulas:**
`=IF($A2="","",IFERROR(INDEX(Sheet1!A:A,MATCH($A2,Sheet1!$E:$E,0)),""))`
`=IF($A2="","",IF(COUNTIF(Sheet1!$E:$E,$A2)=0,"Not found",IF(COUNTIF(Sheet1!$E:$E,$A2)>1,"Duplicate code - first match shown","OK")))`
Filled to row 1001; a drop-down of codes on column A.
**Lessons:**
- Use whole-column references in `INDEX/MATCH` so new products are found without editing formulas.
- A drop-down that reads a named range broke when the file went through Google Sheets (names were dropped). Point the drop-down at a plain range if the file may be re-saved elsewhere.
- Always add a status column ("OK", "Not found", "Duplicate").

## 8. CNIC code to region
**Ask:** When a CNIC such as 54401-45939 is typed, look at the code before the dash; if it matches an earlier record, fill the region.
**Delivered:** `prectice_record_updated.xlsx`, Sheet3. Typing 54401-45939 returns "lahore".
**Formulas:**
- Code: `=IF($G2="","",LEFT(TRIM($G2),IFERROR(FIND("-",TRIM($G2))-1,5)))`
- Known-code helper: `=IF(OR($I2="",$H2="",$H2="New code - type region"),"",$I2)`
- Region (row 7 down): `=IF($G7="","",IFERROR(INDEX(H$1:H6,MATCH($I7,J$1:J6,0)),"New code - type region"))`
**Technique:** the lookup range ends one row above the formula, so it only reads earlier records and can never become circular. Unknown codes show a red "New code" flag; typing a region over it teaches the sheet.
**Lessons:** The user said "6 digit code"; "54401-" is 6 characters including the dash, so the text before the dash (5 digits) was used. State that reading.
Test the edge cases: no dash, a number instead of text, spaces around the CNIC, and a typed override.

## 9. Email validation
**Ask:** Email must end with @buitms.edu.pk; otherwise show "Invalid email" and suggest akbar@buitms.edu.pk.
**Delivered:** `prectice_record_updated_email.xlsx`. Three layers:
1. Data validation on the Email column (custom formula, "stop" alert, input prompt with the example).
2. A visible "email check" column showing OK or "Invalid email - use format akbar@buitms.edu.pk".
3. Conditional formatting (red for invalid, green for OK).
**Formula:** `AND(LEN(TRIM(D2))>14,RIGHT(TRIM(D2),14)="@buitms.edu.pk",ISERROR(FIND(" ",TRIM(D2))),LEN(TRIM(D2))-LEN(SUBSTITUTE(TRIM(D2),"@",""))=1)`
**Lessons:** Data validation only checks what is typed later, never existing or pasted values. That is why the check column exists.
18 test emails were run through both the validation formula and the check column; they agreed on all 18. 7 of the user's current entries were flagged and left untouched.

## 10. Pivot table: monthly sales by region
**Advice given:** Rows = Date grouped by Month (and Year if data spans years); Columns = Region; Values = Sum of Sales; Filters = Product.
**Then built:** Real PivotTables (not formulas) in `sales_with_pivot.xlsx` (new "Pivot" sheet) and `piovt_table_sheet_with_pivot.xlsx` (the empty Sheet2). Grand total 8,156,003 matched the source.
**How:** openpyxl cannot create PivotTables. LibreOffice can, through its scripting interface, and it saves them as real Excel PivotTables. `scripts/make_pivot.py` does this.
**Lessons:**
- All dates were in December 2026, so there was one row only; say so.
- Misspelled regions ("lhaore", "islamabdd") become separate pivot columns; clean the source first.
- A leftover lock file (`.~lock...#`) from a crashed run made the next run fail silently. Delete it before retrying.
- Set `refreshOnLoad` so Excel rebuilds the pivot when opened.

## 11. Shoes Business one-page sales dashboard
**Ask:** One-page dashboard with Sales, Orders, AOV, Margin; suggest layout, charts, slicers, colours; "make sure everything is working".
**Delivered:** `Shoes_Business_Tracker_with_Dashboard.xlsx`: `Dashboard` sheet plus `Dash_Calc` helper sheet. KPI cards with change versus
previous month; Month / Category / Channel drop-down filters; monthly sales + profit chart; channel doughnut; category and top-5 product bars; a self-check line.
**Numbers:** Sales Rs 173,989; 30 orders; AOV Rs 5,800; gross margin 32.7%.
**How it works:** `Dash_Calc` mirrors every Sales row, adds Category from Stock, and sets flags (1/0) for "matches the filters". Every KPI and chart reads from the flags.
Orders count distinct invoices using a first-occurrence flag.
**Testing:** 7 filter combinations compared with an independent calculation (all passed); a shared invoice and a new November sale were also tested.
**Repairs made to the user's file:** the "Dashboard" navigation links pointed to a missing sheet (now fixed); `MAXIFS` was stored without the `_xlfn.` prefix and showed `#NAME?` after a recalculation (201 cells fixed). All original sheets were verified unchanged.
**Lessons:** Real slicers need a PivotTable or Excel Table; drop-down cells that drive `SUMIFS` flags work in every Excel version and in Google Sheets.

## 12. Insights for a sales manager
See `prompt-templates.md`, section 5, for the prompt, the five rules and a worked example from the dashboard numbers above.

---

# Part 3. Prompt templates: how to ask, and how this skill answers

Good Excel answers start with a good question. This file keeps the prompt styles used in the NAVTTC practice
sessions, a cleaned-up version of each, and copy-and-paste templates.

## Contents
1. The six-part prompt (Role, Context, Layout, Task, Format, Example)
2. Prompt library (copy and paste)
3. The original prompts, cleaned up
4. Checklist for a better prompt
5. Insight-writing prompt and a worked example

---

## 1. The six-part prompt

| Part | What to write | Why it matters |
|---|---|---|
| **Role** | "You are an Excel expert" | Sets depth and tone |
| **Context** | What the sheet is and how big it is ("Sales sheet, 2,000 rows") | Tells Claude how far ranges must reach |
| **Layout** | Which column holds what ("A=Date, B=Region, C=Product, D=Sales") | Without it, Claude has to guess column letters |
| **Task** | The one thing you want ("Total sales for North in March") | Keeps the answer focused |
| **Format** | The shape of the answer ("One formula + 2-line explanation") | Controls length and style |
| **Example** | A sample row, a sample answer, or "assume X" | Removes ambiguity |

**Worked example (from the sessions)**

```
Role:    You are an Excel expert
Context: Sales sheet, 2,000 rows
Layout:  A=Date, B=Region, C=Product, D=Sales
Task:    Total sales for North in March
Format:  One formula + 2-line explanation
Example: Assume the region is "North"
```

**Answer that fits the Format line**

```
=SUMIFS(D2:D2001, B2:B2001, "North", A2:A2001, ">="&DATE(2026,3,1), A2:A2001, "<"&DATE(2026,4,1))
```
Adds Sales where Region is "North" and the Date falls from 1 March to 31 March 2026.
Assumes "North", the year 2026, data in rows 2-2001 and real Excel dates (not text).

When the user says "assume X by your own", pick a sensible value, state it in the explanation, and move on. Do not
stop to ask.

---

## 2. Prompt library (copy and paste)

**Formula**
> You are an Excel expert. My sheet has [columns]. Write one formula that [goal]. Put it in [cell]. Give the formula
> and a 2-line explanation. Assume [anything unclear].

**Update my file**
> Use the attached file. Add [feature] to [sheet/column]. Keep my existing data and formatting. Give me the updated file
> and tell me what you changed.

**Validation rule**
> In the [column] of [sheet], accept only values that [rule]. Otherwise show "[message]" and suggest this format:
> [example]. Give me the updated sheet.

**Lookup (type a code, get the whole row)**
> Sheet 1 holds the master data. In Sheet 2, when I type a [code] in column A, fill in the rest of the row from
> Sheet 1. It must work for every row, and show "Not found" for unknown codes.

**Duplicates and blanks**
> Read every row and column. Count each category in [column]. Show how many duplicate values there are in [column].
> Show how many unique and duplicate values there are in [columns]. Count empty cells in the whole sheet. Then give a
> complete analysis with charts.

**Pivot table**
> I have [columns]. Design a pivot table to compare [measure] by [dimension] per [period]. Tell me what goes in Rows,
> Columns, Values and Filters. Then build it in my file.

**Dashboard**
> Design a one-page Excel dashboard for a manager. KPIs: [list]. Suggest layout, charts, filters and colours. Build it in
> my file and make sure everything works.

**Fix or audit**
> My formula `[paste]` returns [error or wrong result]. The data looks like [sample rows]. Find the cause and give a
> corrected formula.

**Insights for a manager** (see section 5)

---

## 3. The original prompts, cleaned up

| # | What was asked (typos fixed) | What it taught |
|---|---|---|
| 1 | "Read each row and column. Count each category by Branch Name. Find how many duplicate names are in the Name column. Show unique and duplicate Addresses and identity numbers. Count null cells in the whole sheet. Generate the complete analysis with relevant graphs." | Say exactly which columns; "identity" was read as CNIC, and the reading was stated back |
| 2 | "Write an Excel formula that returns Bonus if sales are 50,000, else No bonus." | "Are 50,000" is ambiguous: `>=` was used and `=` offered |
| 3 | "Give me the updated sheet with the formula applied." | Deliver the file, not just the formula |
| 4 | Six-part prompt: total sales for North in March (assume North) | Use the Format line to control length |
| 5 | "Fetch the sales greater than 50,000 with all date, region, product, then update the sheet." | Extract matching rows with formulas so the result updates itself |
| 6 | "Highlight the highest sales, then make the graph of all sales according to sales values." | Highlight with conditional formatting, not by hand |
| 7 | "Make the graph not simple." | Combo charts, KPI tiles and live insights beat a basic column chart |
| 8 | "Two sheets in one book. When I enter a product code in Sheet 2 it shows the complete row. Do it for all products." | INDEX/MATCH lookup plus a status column |
| 9 | "When someone enters a CNIC, look at the code before the dash. If it matches an earlier record, fill the region." | Learn from earlier rows; flag unknown codes |
| 10 | "Email must end with @buitms.edu.pk, else show invalid email and suggest akbar@buitms.edu.pk." | Validation rule plus a visible check column |
| 11 | "Design a pivot table to compare monthly sales by region. What goes in Rows, Columns, Values, Filters?" | Answer first, then offer to build it |
| 12 | "Another sheet. Make the pivot table in both files on your own." | Build real PivotTables, not just a description |
| 13 | "Design a one-page sales dashboard: Sales, Orders, AOV, Margin. Suggest layout, charts, slicers, colours. Make sure everything works." | "Make sure it works" means test it, not just build it |
| 14 | "Summary table: [paste table]. Write 5 bullet insights for a sales manager. Exact numbers, biggest change, one action. No jargon." | A reusable report prompt (section 5) |

---

## 4. Checklist for a better prompt

- Name the file and the sheet, or attach it.
- Say which column holds what (letters or header names).
- Say how many rows (or "will grow").
- Give the exact rule, with one example that should pass and one that should fail.
- Say what the output should look like (one formula, a new sheet, a chart, a file).
- Say what must not change ("keep my data").
- Mention your Excel version if you use Excel 2016 or older (affects FILTER, XLOOKUP, dynamic arrays).

---

## 5. Insight-writing prompt and a worked example

**Prompt**
> Here is a summary table: [paste table]. Write 5 bullet insights for a sales manager. Use exact numbers, mention the
> biggest change, and suggest one action. No jargon.

If the table is missing (the placeholder `[paste table]` is still there), say so and ask for it. Do not invent numbers.

**Rules the answer follows**
1. Exactly 5 bullets, each one sentence or two short ones.
2. Every bullet has at least one exact number from the table (Rs 173,989, not "about 174k").
3. One bullet names the biggest change: the metric, the two values, the gap in rupees and in percent.
4. After the bullets, one line starting "Action:" with one concrete step.
5. Plain words: "sales" not "revenue run-rate", "fell" not "contracted", "share" not "mix".

**Worked example** (numbers from the Shoes Business Tracker dashboard, July to October 2026)

- Sales were Rs 173,989 from 30 orders, so the average order was Rs 5,800.
- Biggest change: September sales fell 85.9% from August (Rs 48,198 to Rs 6,799, a drop of Rs 41,399); October then
  recovered to Rs 39,996, up 488.3%.
- Oxford Leather alone brought in Rs 70,000, which is 40.2% of all sales; the top three products (Oxford Leather,
  Street Sneaker, Winter Boot) bring 82.4%.
- Sales are spread across channels: Shop leads with Rs 44,297 (25.5%) and Instagram is lowest with Rs 24,300 (14.0%).
- Profit stayed near a third of sales every month, but slipped from 33.4% in July to 31.5% in October.

**Action:** No sales are recorded after 20 October, so confirm November sales are being entered, then push the three best
sellers on Instagram, where sales are lowest.

---

# Part 4. Formula cookbook

Most formulas here were built and tested in the NAVTTC sessions. Rows marked *(standard)* are well-known patterns added
from general Excel practice.

Row 1 is the header and data starts in row 2 unless stated. Change the ranges to match the sheet.

## Contents
1. Conditions and flags
2. Sums, counts and averages by condition
3. Lookups
4. Extracting rows (no FILTER function)
5. Ranking and sorting
6. Duplicates, unique values and blanks
7. Text tricks (codes, emails, cleaning)
8. Data validation rules
9. Conditional formatting rules
10. KPI and dashboard formulas
11. Compatibility notes

---

## 1. Conditions and flags
| Need | Formula |
|---|---|
| Bonus at or above a target | `=IF(A2>=50000,"Bonus","No bonus")` |
| Blank-safe result | `=IF(A2="","",your_formula)` |
| Only catch expected misses | `=IFERROR(INDEX(...),"Not found")` |
| 1/0 flag for several conditions | `=IF(AND(B2="North",D2>0),1,0)` |
| "All" option in a filter | `=IF(OR($T$7="All categories",F5=$T$7),1,0)` |

Tip: use `>=` or `>` deliberately and say which one you chose.

## 2. Sums, counts and averages by condition
| Need | Formula |
|---|---|
| Sales for a region in a month | `=SUMIFS(D2:D2001,B2:B2001,"North",A2:A2001,">="&DATE(2026,3,1),A2:A2001,"<"&DATE(2026,4,1))` |
| Count rows above a value | `=COUNTIF(D2:D2001,">"&F5)` |
| Count with two conditions *(standard)* | `=COUNTIFS(B2:B2001,"North",D2:D2001,">50000")` |
| Average with a condition *(standard)* | `=AVERAGEIFS(D2:D2001,B2:B2001,"North")` |
| Blank cells in a range | `=COUNTBLANK(A3:X2514)` |
| Blank cells in one row | `=COUNTBLANK(NIB!A3:X3)` |
| Cells in a region of the sheet | `=ROWS(NIB!A1:Y2)*COLUMNS(NIB!A1:Y2)` |
| Text "NA" with stray spaces | `=SUMPRODUCT(--(UPPER(TRIM(G3:G2514))="NA"))` |
| Month label | `=TEXT(A2,"mmm yyyy")` |
| First day of the month | `=DATE(YEAR(A2),MONTH(A2),1)` |
| Previous month | `=EDATE(T9,-1)` |

Always use a date range (`>=` first day, `<` first day of next month) rather than `MONTH(A:A)=3`; it is faster, needs no
array formula, and does not mix years.

## 3. Lookups
| Need | Formula |
|---|---|
| Return a column for a code | `=IF($A2="","",IFERROR(INDEX(Sheet1!A:A,MATCH($A2,Sheet1!$E:$E,0)),""))` |
| Status next to a lookup | `=IF($A2="","",IF(COUNTIF(Sheet1!$E:$E,$A2)=0,"Not found",IF(COUNTIF(Sheet1!$E:$E,$A2)>1,"Duplicate code - first match shown","OK")))` |
| Two-way lookup *(standard)* | `=INDEX(B2:M20,MATCH(P1,A2:A20,0),MATCH(P2,B1:M1,0))` |
| Category from another sheet | `=IFERROR(INDEX(Stock!$C$5:$C$104,MATCH($E5,Stock!$A$5:$A$104,0))&"","Unknown")` |
| Learn from earlier rows only | `=IF($G7="","",IFERROR(INDEX(H$1:H6,MATCH($I7,J$1:J6,0)),"New code - type region"))` |

Why `INDEX/MATCH` and not `VLOOKUP`: it can return a column to the left of the key, it does not break when columns are
inserted, and whole-column references (`Sheet1!$E:$E`) make new rows work automatically. Use `XLOOKUP` only if the user is
on Excel 365 and says so.

## 4. Extracting rows (no FILTER function)
Helper column on the data sheet (E) numbers matching rows:
```
=IF(AND(ISNUMBER(D2),D2>High_Sales!$B$3),COUNTIF(D$2:D2,">"&High_Sales!$B$3),"")
```
Result sheet: threshold in B3, number of matches in B4, and in row 7 down:
```
A7: =IF(ROWS(A$7:A7)<=$B$4,ROWS(A$7:A7),"")
B7: =IF($A7="","",INDEX(Sheet1!A$2:A$2001,MATCH($A7,Sheet1!$E$2:$E$2001,0)))
B4: =COUNTIF(Sheet1!$D$2:$D$2001,">"&$B$3)
```
Fill B7 across for each column. Add a warning if matches exceed the number of prepared rows.

## 5. Ranking and sorting
Unique rank with tie-break (ties keep their original order):
```
=RANK(I6,$I$6:$I$17)+COUNTIF($I$6:I6,I6)-1
```
Sorted list: `=INDEX(Sheet1!A$2:A$13,MATCH($A6,$J$6:$J$17,0))` where A6 is 1, 2, 3 ...
Top 5: `=IFERROR(INDEX($AB$5:$AB$104,MATCH(ROWS(AS$5:AS5),$AE$5:$AE$104,0)),"")`
Cumulative share (Pareto): `=SUM(B$49:B49)/SUM(B$49:B$60)`
Running total: `=SUM(F$49:F49)`

## 6. Duplicates, unique values and blanks
Build a clean key first, then count on the key:
```
Key (text):     =IF(NIB!F3="","",UPPER(TRIM(NIB!F3)))
ID key:         =IF(NIB!E3="","",UPPER(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(NIB!E3&"",".",""),"-","")," ",""))))
Count of key:   =IF(C3="","",SUMPRODUCT(--($C$3:$C$2514=C3)))
First seen:     =IF(C3="","",IF(MATCH(C3,$C$3:$C$2514,0)=ROW()-2,1,0))
Status:         =IF(C3="","Blank",IF(D3=1,"Unique","Duplicate"))
Placeholder:    =IF(K3="","Blank",IF(OR(K3="NA",K3="ADDRESS MISSING"),"Placeholder",IF(L3=1,"Unique","Duplicate")))
```
Metrics from the status columns:
| Metric | Formula |
|---|---|
| Records that are duplicates | `=COUNTIF(F3:F2514,"Duplicate")` |
| Distinct values that repeat | `=COUNTIFS(E3:E2514,1,F3:F2514,"Duplicate")` |
| Distinct values | `=SUM(E3:E2514)` |
| Unique (appear once) | `=COUNTIF(F3:F2514,"Unique")` |
| Extra repeat rows | `=records_in_duplicates - distinct_repeating_values` |
| Check the arithmetic | `=IF(AND(C10+C11=C9,C6+C7+C10+C12=C5),"OK","CHECK")` |

(`ROW()-2` assumes the data starts in row 3; adjust to `ROW()-first_row+1`.)

Define terms in the answer: "unique" (appears exactly once), "distinct" (each value counted once), "duplicate value"
(a value that repeats), "records in duplicates" (all the rows that share a repeated value).

Why `SUMPRODUCT(--(range=cell))` and not `COUNTIF`: `COUNTIF` treats `* ? ~` as wildcards, treats text that looks like a
number as a number, and stops at 15 digits. `SUMPRODUCT` compares exactly.

## 7. Text tricks (codes, emails, cleaning)
| Need | Formula |
|---|---|
| Code before the dash (fallback: first 5) | `=IF($G2="","",LEFT(TRIM($G2),IFERROR(FIND("-",TRIM($G2))-1,5)))` |
| Remove all spaces *(standard)* | `=SUBSTITUTE(A2," ","")` |
| Non-breaking space | `=SUBSTITUTE(A2,CHAR(160)," ")` (TRIM does not remove it) |
| Ends with a domain | `=RIGHT(TRIM(D2),14)="@buitms.edu.pk"` (14 = length of the domain) |
| Count of "@" | `=LEN(D2)-LEN(SUBSTITUTE(D2,"@",""))` |
| Number to text without formats | `=A2&""` |
| Join with a label | `=TRIM(C6)&" - "&TRIM(D6)` |

## 8. Data validation rules
**Email must end with a domain** (Data > Data Validation > Custom; formula relative to the first cell):
```
=AND(LEN(TRIM(D2))>14,RIGHT(TRIM(D2),14)="@buitms.edu.pk",ISERROR(FIND(" ",TRIM(D2))),LEN(TRIM(D2))-LEN(SUBSTITUTE(TRIM(D2),"@",""))=1)
```
Settings: Error Alert style Stop, title "Invalid email", message "Invalid email. Please use this format: akbar@buitms.edu.pk".
Input Message: "Use your university email, for example: akbar@buitms.edu.pk".
Add a check column so old and pasted values are flagged too:
```
=IF(TRIM($D2)="","",IF(<same formula>,"OK","Invalid email - use format akbar@buitms.edu.pk"))
```
**Drop-down list:** Allow List, Source `=Sheet1!$E$2:$E$100` (a plain range survives re-saving in other tools better than a named range).
**Number between** *(standard)*: Allow Whole number, between 1 and 100. **Date** *(standard)*: Allow Date, greater than 01-Jan-2026.
Data validation never re-checks cells that already hold data, and pasting over a cell bypasses it.

## 9. Conditional formatting rules
| Need | Rule type and formula |
|---|---|
| Highlight the highest sale (whole row) | Formula, applies to `A2:D2001`: `=AND(ISNUMBER($D2),$D2=MAX($D$2:$D$2001))` |
| Flag a new code | Applies to `H7:H1001`: `=$H7="New code - type region"` (red fill) |
| Invalid email | Applies to `D2:D1001`: `=AND($D2<>"",$K2<>"OK")` (red) |
| Green up arrow, red down arrow | `=LEFT($B$9,1)="▲"` (green) and `=LEFT($B$9,1)="▼"` (red) |
| Duplicates *(standard)* | Built-in Duplicate Values, or `=COUNTIF($A$2:$A$2001,$A2)>1` |

## 10. KPI and dashboard formulas
| KPI | Formula |
|---|---|
| Sales | `=SUMIFS(Revenue,Flag,1)` |
| Orders (distinct invoices) | flag per row: `=IF($K5=1,IF($D5="",1,IF(COUNTIFS($D$5:$D5,$D5,$K$5:$K5,1)=1,1,0)),0)`; then `=SUM(flag column)` |
| Average order value | `=IF(Orders=0,0,Sales/Orders)` |
| Gross margin % | `=IF(Sales=0,0,Profit/Sales)` |
| Change versus previous month | `=IF(prev=0,"no sales in "&label,IF(cur>=prev,"▲ ","▼ ")&TEXT(ABS(cur/prev-1),"0.0%")&" vs "&label)` |
| Margin change in points | `=TEXT(ABS(cur-prev)*100,"0.0")&" pts"` |
| Share of total *(standard)* | `=B2/SUM($B$2:$B$9)` |
| Month filter flag | `=IF(OR($T$6="All months",$C5=$T$6),1,0)` |
| Rolling 6-month start | `=MAX(first_month,EDATE(last_month,-5))` |

## 11. Compatibility notes
- Prefer: `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `IFERROR`, `SUMPRODUCT`, `TEXT`, `EDATE`. They work in Excel 2007+, LibreOffice and Google Sheets.
- Avoid unless the user is on Excel 365: `FILTER`, `UNIQUE`, `SORT`, `XLOOKUP`, `LET`, spilled ranges.
- When writing files with code, functions added after Excel 2007 must be stored with a prefix or they show `#NAME?` after a
  recalculation: `_xlfn.MAXIFS`, `_xlfn.MINIFS`, `_xlfn.IFS`, `_xlfn.SWITCH`, `_xlfn.TEXTJOIN`, `_xlfn.CONCAT`, `_xlfn.XLOOKUP`.
  Typing them in Excel adds the prefix automatically.
- Whole-column references are fine in `INDEX/MATCH/SUMIFS/COUNTIF`; avoid them inside `SUMPRODUCT` and array formulas (slow).

---

# Part 5. Dashboards and pivot tables

## Contents
1. Pivot table design (what goes where)
2. Building a real PivotTable from code
3. One-page dashboard design
4. The calc-sheet pattern (filters that really work)
5. Chart choices and colours
6. Slicers: real versus drop-down filters
7. Print and delivery checklist

---

## 1. Pivot table design (what goes where)

| Area | Put here | Example: compare monthly sales by region |
|---|---|---|
| **Rows** | What you read down the page (usually time or a category) | Date, grouped by Month (add Years if data spans years) |
| **Columns** | What you compare side by side | Region |
| **Values** | The number being measured, with the right summary | Sum of Sales (use Count only for "how many") |
| **Filters** | A way to slice the whole table without changing the layout | Product (and Year, if grouped) |

Steps in Excel: click inside the data > Insert > PivotTable > New Worksheet > drag the fields > right-click a date > Group >
Months (and Years) > Value Field Settings > Number Format.

Rules of thumb:
- If there are many columns (more than about 8), swap Rows and Columns so the table does not run off the screen.
- If Values shows "Count of Sales", the numbers are stored as text; fix the source.
- Misspelled categories ("lhaore" vs "Lahore") appear as separate rows or columns. Clean the source first.
- A pivot does not update by itself: right-click > Refresh. If the data grows, use Change Data Source, or base the pivot on an Excel
  Table (Ctrl+T) so it grows automatically.
- Grouping by Months only merges the same month of different years. Add Years.
- Say honestly when the data has only one month: a month-by-month comparison needs more than one month.

## 2. Building a real PivotTable from code
`openpyxl` cannot create PivotTables. LibreOffice can, through its scripting interface, and it saves real Excel PivotTables.
Use `scripts/make_pivot.py`:

```
python scripts/make_pivot.py --src sales.xlsx --out sales_with_pivot.xlsx --source-sheet Sheet1 \
       --target-sheet Pivot --rows date --cols region --pages product --data sales:sum --group date
```
What it does: finds the used range, creates the pivot at A5 of the target sheet (creating the sheet if needed), groups the date field
by month (`--group-by quarters|years|days` for other groupings), saves as .xlsx, then patches the file so Excel refreshes the pivot when it opens and so labels read "Sum of sales" and
"Grand Total".

Known limits:
- Grouping offers one level at a time (months, quarters or years), not months plus years together.
- The pivot reads a fixed range; after adding rows use Change Data Source.
- It could not be opened in desktop Excel in the sandbox. It was verified by re-opening in LibreOffice, parsing the XML, and checking totals.
- Delete a leftover `.~lock.<file>#` before re-running if a previous run crashed.

## 3. One-page dashboard design

**Layout (landscape, one page):**
1. Header band: title, the active filters in words, and "data up to <date>".
2. Filter row: Month, Category, Channel (drop-down cells).
3. KPI cards (4): Sales, Orders, Average order value, Margin. Each shows a change versus the previous period.
4. Row of two charts: trend over time (wide) and share by channel (narrow).
5. Row of two charts: ranking by category and top products.
6. Footer: a live data check ("every sale is included") and the definitions of the KPIs.

**KPI definitions to print on the page:** Sales = revenue after discounts. Orders = distinct invoices. AOV = Sales / Orders.
Gross margin = profit / Sales (cost of goods only, before marketing and other expenses).

**Charts that answer manager questions:**
| Question | Chart |
|---|---|
| Is it going up or down? | Columns for sales plus a line for profit, by month |
| Where do sales come from? | Doughnut or bar by channel |
| What sells? | Sorted bars by category and top 5 products |
| How concentrated is it? | Pareto (columns plus cumulative % line on a second axis) |
| What is the typical sale? | KPI tiles for median next to average |

If one value is far larger than the rest (one sale was 97% of the total), show a second view: log scale, or "everything
except the top item".

**Colours (matching a navy workbook):** navy `1F3864` for headers and Sales, teal `0F7B8A` for Orders, purple `5B4B9A` for AOV,
green `2E8B57` for Margin, blue `2E75B6` for columns, amber `F39C12` for the profit line and highlights, grey `9AA0A6` for "other".
Green for up and red for down. Blue text means "type or click here". Keep to 5 or 6 colours and use each for one meaning.

## 4. The calc-sheet pattern (filters that really work)
Put all logic on a helper sheet (`Dash_Calc`) so the dashboard sheet holds only labels, KPI cells and charts.

1. **Filter state:** read the three drop-down cells and turn blanks into "All ..." values.
2. **Row mirror:** one row per source record: date, month label, invoice, product, category (looked up), channel, revenue, profit.
3. **Flags (1/0):** `Match: category + channel` and `Match: all filters` (adds the month). A third flag marks the previous month.
4. **KPIs** are `SUMIFS(value, flag, 1)`. Orders use a first-occurrence flag on the invoice so shared invoices count once.
5. **Chart tables:** months window, sorted categories, top 5, channels. Charts read these ranges only.
6. **Checks:** category total = Sales, channel total = Sales, rows read. Show one summary line on the dashboard.
7. **Drop-down lists:** a month list built with `EDATE`, a category list built from a first-occurrence flag, a fixed channel list.

Why: charts on top of formulas that read the flags update the moment a filter changes, with no macros and no pivot refresh.

## 5. Chart choices and tips
- Use native Excel charts so the user can edit them. Set titles, axis titles and data labels; do not rely on colour alone.
- For a horizontal ranking, reverse the category axis so rank 1 is at the top.
- Combo charts (columns + line) share one axis when both are the same unit; put percentages on a second axis.
- Do not use #N/A to make gaps in a line; use zeros for amounts, or leave margin out of the trend.
- Check the picture: render the page to an image and look at it.

## 6. Slicers: real versus drop-down filters
- **Real Excel slicers** need a PivotTable or an Excel Table. They are made in Excel (Insert > Slicer) and cannot be reliably
  created by code. Use them when the dashboard is built on pivots.
- **Drop-down filter cells** (data validation list) drive `SUMIFS`/flag formulas, work in every Excel version and in Google Sheets,
  and can be tested by code. Use these for formula-driven dashboards.
- Always add an "All ..." option and a message when a filter combination has no sales.

## 7. Print and delivery checklist
- Landscape, fit to 1 page wide by 1 page tall, print area set, narrow margins, gridlines off.
- Hide nothing; leave the helper sheet visible and say "do not delete".
- Repair navigation links that point to the dashboard.
- Verify: zero formula errors; numbers match an independent calculation; at least 5 filter combinations (including one with no data); render and look.

---

# Part 6. Data quality, pitfalls and verification

## Contents
1. Profile the file before you build anything
2. Blank, placeholder and zero are three different things
3. Pitfalls learned the hard way
4. Files that pass through Google Sheets
5. Verification protocol ("make sure everything works")
6. Describing results to a non-technical reader

---

## 1. Profile the file before you build anything
Run `scripts/profile_data.py <file> --header-row N` and read the result, or check by hand:
- **Sheets and size:** names, used range, hidden sheets, merged cells, frozen panes, existing validation and conditional formats.
- **Header row:** is it row 1? A title above the headers (NIB had a title in row 1 and headers in row 2) shifts every range.
- **Footer or total rows:** rows below the data with totals or labels. Exclude them from record counts and say so.
- **Empty trailing columns:** a stray empty column at the end is common.
- **Data types per column:** numbers stored as text, dates stored as text, mixed columns (numeric and text IDs).
- **Blanks versus placeholders:** see section 2.
- **Duplicates:** by exact value and after cleaning (case, extra spaces, dashes).
- **Which cells are formulas:** helper columns, auto-filled columns and existing formulas must not be overwritten.
- **Existing quirks:** array formulas, broken links, text that looks like numbers.

State your reading of the request in one line ("'identity' taken to mean the CNIC column") before presenting results.

## 2. Blank, placeholder and zero are three different things
| Looks like | Really is | How to count |
|---|---|---|
| Empty cell | Null / missing | `COUNTBLANK(range)` |
| `NA`, `N/A`, `NIL`, `-`, `ADDRESS MISSING` | Text that means "missing" (placeholder) | `SUMPRODUCT(--(UPPER(TRIM(range))="NA"))` |
| `0` | A real zero | `COUNTIF(range,0)` |
| Spaces only | Looks blank, is not | `SUMPRODUCT(--(LEN(TRIM(range))=0),--(range<>""))` |

Report them separately. In the NIB file 5 addresses were truly empty and 337 held placeholder text; counting only blanks would have hidden 98% of the problem.

Also separate: blanks that are **by design** (an account number on a payment-instrument record, FX-rate fields on non-FX
records, optional notes) from **real gaps** (a missing CNIC or name). Label each column.

## 3. Pitfalls learned the hard way
| Pitfall | What happens | Fix |
|---|---|---|
| Reading with pandas | `NA`, `N/A`, `NULL` become empty values silently | Read with `openpyxl`, or `keep_default_na=False` |
| Blank turned into text | A missing value becomes the text `nan` and creates a false duplicate | Test with `pd.isna()` before converting; compare counts to an independent calculation |
| `COUNTIF` on keys | `* ? ~` act as wildcards; numbers stored as text match numbers; 15-digit limit | `SUMPRODUCT(--(range=cell))` |
| `MATCH` with text | Same wildcard behaviour | Check that keys contain none, or use `SUMPRODUCT` |
| Case and spaces | "Ali" and "ali " count as different | Build a key with `UPPER(TRIM())` |
| `TRIM` does not remove non-breaking spaces | Keys still differ | `SUBSTITUTE(x,CHAR(160)," ")` |
| ID formats | Dotted, dashed and plain versions of one ID differ | Remove `. - space` in the key |
| Footer rows | Totals counted as records | Limit ranges to the data rows |
| Text dates | `SUMIFS` with date criteria returns 0 | Convert to real dates; check with `ISNUMBER` |
| Day-first regional settings | `1/13/2026` cannot be a date, stays text | Mention it, or write dates as real dates |
| Dominant outlier | One value hides the rest in a chart | Second chart without it, or log scale |
| `IFERROR` everywhere | Hides real mistakes | Use it only where a miss is expected |
| Hard-coded results | Go stale | Use formulas; keep assumptions in labelled cells |
| Newer functions written by code | `#NAME?` after a recalculation | Use the `_xlfn.` prefix, or avoid them |
| Many formulas with growing ranges | Slow | Fine up to a few thousand rows; avoid whole-column arrays |
| A crashed LibreOffice run | A `.~lock.file#` blocks the next run, which then returns nothing | Delete lock files, kill stray `soffice.bin` |
| `pkill -f soffice` | Kills your own shell command | `pkill -9 soffice.bin` |
| CSV | Cannot store formulas, formats or charts | Deliver .xlsx and say why |
| Overwriting a user's data | Trust lost | Add new sheets or columns; prove with a before/after diff |

## 4. Files that pass through Google Sheets
Several uploads were exported from Google Sheets. Expect:
- **Lost defined names.** A drop-down that read a named range stopped working. Point validation lists at plain ranges.
- **Array formulas** (`{=...}`) written for ordinary formulas. They still work; preserve them.
- **Broken internal links** such as `null!A1` when the target sheet did not exist.
- **Unprefixed newer functions** such as `MAXIFS`, which Excel and LibreOffice show as `#NAME?` after a recalculation. Add `_xlfn.`.
- **No document properties and no cached styles** from Excel. Do not assume Excel-only features survive.
Formulas you write should work in both Excel and Google Sheets: `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `IFERROR`, `TEXT`, `EDATE`.

## 5. Verification protocol ("make sure everything works")
A recalculation with zero errors proves only that nothing crashed. Do all of these:
1. **Recalculate** (LibreOffice headless) and confirm zero formula errors. Investigate every error and say whether it was already in the user's file.
2. **Independent check:** recompute the key numbers in Python (or by hand) from the raw inputs and compare. Match every KPI and chart table.
3. **Scenario tests:** change the inputs and re-check. For a filtered dashboard: all filters off; one month; a month with no earlier month; two filters
   together; a selection with no data; a shared invoice; a newly added row.
4. **Edge cases for lookups and validation:** blank input, unknown code, extra spaces, lower case, number instead of text, wrong format.
5. **Original data untouched:** compare every original cell (values and formulas) before and after; also validation, conditional formats, freeze panes.
6. **Look at it:** convert to PDF and view the image of each new sheet. Check one page, labels readable, nothing overlapping.
7. **Report honestly:** what was tested, what could not be tested (for example "not opened in desktop Excel"), and every repair or assumption.

## 6. Describing results to a non-technical reader
- Lead with the answer (the formula, the file, the headline number), then the caveats.
- Use exact numbers with units (Rs 173,989; 32.7%), not "about".
- Define words that could mean two things: unique, distinct, duplicate, null, blank.
- Name problems in the data plainly (misspelled regions, placeholder addresses) and say what they affect.
- Use plain words. No jargon. Offer the next step in one sentence.

---

# Part 7. Excel expertise notes (general knowledge for a data analyst)

These notes go beyond the sessions. They are the standard toolkit an Excel analyst should have ready.

## Contents
1. Choosing the right tool
2. Function families to know
3. Error messages and what they mean
4. Speed and good structure
5. Auditing and debugging
6. Cleaning data
7. Statistics and business maths
8. Dates and times
9. Formatting that helps readers
10. Shortcuts worth learning
11. Power Query, Power Pivot and macros (when formulas stop being enough)
12. A note on identity numbers (CNIC)

---

## 1. Choosing the right tool
| Situation | Use |
|---|---|
| One number from conditions | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| Look something up | `INDEX/MATCH` (any version), `XLOOKUP` (365) |
| Summarise many rows by categories | PivotTable |
| Live, filter-driven report | Formula flags + drop-downs, or pivots + slicers |
| Prevent wrong entry | Data validation (plus a check column) |
| Show a rule visually | Conditional formatting |
| Clean or merge messy files repeatedly | Power Query |
| Millions of rows or many related tables | Power Pivot (data model) |
| Repetitive clicking | Macro (VBA) or Office Scripts |

## 2. Function families to know
- **Logical:** `IF`, `IFS`, `AND`, `OR`, `NOT`, `IFERROR`, `IFNA`, `SWITCH`.
- **Lookup:** `INDEX`, `MATCH`, `VLOOKUP` (exact match needs 0/FALSE), `HLOOKUP`, `XLOOKUP`, `LOOKUP`, `CHOOSE`, `OFFSET` (volatile, avoid).
- **Math and stats:** `SUM`, `SUMIFS`, `SUMPRODUCT`, `ROUND`, `ROUNDUP`, `MOD`, `AVERAGE`, `MEDIAN`, `MODE.SNGL`, `LARGE`, `SMALL`, `RANK.EQ`, `PERCENTILE.INC`, `STDEV.S`, `CORREL`.
- **Text:** `LEFT`, `MID`, `RIGHT`, `FIND` (case-sensitive), `SEARCH` (not), `LEN`, `TRIM`, `CLEAN`, `UPPER`, `LOWER`, `PROPER`, `SUBSTITUTE`, `REPLACE`, `TEXT`, `TEXTJOIN`, `VALUE`.
- **Date:** `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `EDATE`, `EOMONTH`, `WEEKDAY`, `WEEKNUM`, `NETWORKDAYS`, `DATEDIF`.
- **Information:** `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISERROR`, `COUNTA`, `COUNT`, `COUNTBLANK`, `ROW`, `ROWS`, `COLUMN`.
- **Financial:** `PMT`, `FV`, `PV`, `NPV`, `IRR`, `RATE`.

`VLOOKUP` versus `INDEX/MATCH` versus `XLOOKUP`: `VLOOKUP` only looks to the right of the key and breaks when a column is inserted. `INDEX/MATCH`
has no such limit. `XLOOKUP` is the neatest, but only Excel 365/2021 has it.

Approximate-match bands (grades, tax slabs, commission tiers): `=LOOKUP(B2,{0,40,50,60,70,80},{"F","D","C","B","A-","A"})` or a small
table with `=INDEX(rates,MATCH(B2,limits,1))` (limits sorted ascending).

## 3. Error messages and what they mean
| Error | Usual cause | Fix |
|---|---|---|
| `#N/A` | Lookup value not found | Check spelling, spaces, text-versus-number; use `IFNA`/`IFERROR` only if a miss is normal |
| `#VALUE!` | Wrong type (text in a calculation, a range where one value is expected) | Check with `ISNUMBER`; clean with `VALUE`/`TRIM` |
| `#REF!` | A referenced cell or sheet was deleted | Re-point the reference |
| `#DIV/0!` | Dividing by zero or blank | `=IF(B2=0,0,A2/B2)` |
| `#NAME?` | Misspelt function or name; newer function in an old version | Fix the spelling; add `_xlfn.` when writing files by code |
| `#NUM!` | Impossible number (square root of a negative) | Check inputs |
| `#SPILL!` | A dynamic array has no room to spread | Clear the cells below |
| `#####` | Column too narrow, or a negative date | Widen the column |

## 4. Speed and good structure
- **Keep inputs, calculations and outputs apart.** Inputs on a data sheet; helper columns on a calc sheet; the report on its own sheet.
- **Do not hard-code.** Put thresholds (50,000), names ("North") and rates in labelled cells and refer to them.
- **Absolute versus relative references:** `$A$1` fixed, `A$1` fixed row, `$A1` fixed column. `F4` cycles through them.
- **Excel Tables (Ctrl+T)** grow automatically; formulas use names such as `Sales[Amount]`; pivots and charts built on a table follow it.
- **Volatile functions** (`TODAY`, `NOW`, `RAND`, `OFFSET`, `INDIRECT`) recalculate constantly. Use sparingly.
- **Whole-column references** are fine in `SUMIFS`, `COUNTIF`, `INDEX/MATCH`; avoid inside `SUMPRODUCT` and array formulas.
- **Avoid merged cells in data areas** (they break sorting, filtering and pivots); use "Center Across Selection" or keep merges for report headers only.
- **One header row, one record per row, no blank rows or columns** inside a data set.
- **Name sheets clearly** and avoid spaces if formulas will refer to them often.

## 5. Auditing and debugging
- **Formulas > Show Formulas** (Ctrl+`) to see all formulas at once.
- **Evaluate Formula** steps through a calculation.
- **Trace Precedents / Dependents** draws arrows to the cells a formula uses or feeds.
- **F9** on a selected part of a formula shows its value (press Esc to cancel, never Enter).
- **Check totals two ways:** `SUM` of a column against the sum of a pivot; counts against `ROWS`.
- **Add check cells** that say "OK" or "CHECK" next to important totals.
- **Go To Special** (F5 > Special) selects blanks, formulas, constants, errors.

## 6. Cleaning data
- `TRIM`, `CLEAN`, `SUBSTITUTE(x,CHAR(160)," ")` for stray spaces and non-printing characters.
- Text to Columns (Data tab) splits "ID-Name" or fixes text dates and numbers stored as text.
- Flash Fill (Ctrl+E) for quick pattern extraction.
- Remove Duplicates (Data tab) deletes rows, so copy first, or prefer a flag column with `COUNTIF`/`SUMPRODUCT`.
- Fix inconsistent categories with a lookup table (`spelling` > `standard name`) instead of editing by hand.
- Convert text to numbers: `=VALUE(A2)` or multiply by 1. Convert numbers to text: `=A2&""`.
- Standardise phone numbers, IDs and dates before comparing or counting.

## 7. Statistics and business maths
- **Percent change:** `=(new/old)-1`. **Percent of total:** `=B2/SUM($B$2:$B$9)`. **Margin:** `=profit/sales`. **Markup:** `=(price-cost)/cost`.
- **Average order value:** sales / orders (count orders as distinct invoices).
- **Median versus average:** if they differ a lot, outliers are pulling the average. Report both.
- **Moving average:** `=AVERAGE(B2:B4)` filled down. **Running total:** `=SUM(B$2:B2)`.
- **Compound growth (CAGR):** `=(end/start)^(1/years)-1`.
- **Pareto (80/20):** sort descending, then cumulative share. Useful for products and customers.
- **Trend line:** add a chart trendline, or `=FORECAST.LINEAR(x,known_y,known_x)`.
- **Outliers:** values beyond `Q3 + 1.5 x (Q3 - Q1)`, with `QUARTILE.INC`.
- **Correlation:** `=CORREL(range1,range2)`. Correlation is not proof of cause.

## 8. Dates and times
- Dates are numbers (days since 1900); format decides how they look. Times are fractions of a day.
- A text date will not work in `SUMIFS`. Test with `ISNUMBER(A2)`.
- First of month `=DATE(YEAR(A2),MONTH(A2),1)`; end of month `=EOMONTH(A2,0)`; next month `=EDATE(A2,1)`.
- Whole-month filters: `>= first day` and `< first day of the next month`.
- Age or tenure: `=DATEDIF(A2,TODAY(),"y")`. Working days: `=NETWORKDAYS(A2,B2)`.
- `TEXT(A2,"mmm yyyy")` gives labels such as "Jul 2026" (month names depend on the language of Excel).

## 9. Formatting that helps readers
- Blue text for inputs, black for formulas; one header style; consistent number formats.
- Number formats: `#,##0` (thousands), `"Rs "#,##0` (currency label), `0.0%` (percent), `dd-mmm-yyyy` (date), `mmm yyyy` (month).
- Freeze panes under the header; filters on the header row; sensible column widths; no gridlines on reports.
- Conditional formatting for exceptions (red for errors, green for OK), not for decoration.
- Keep a legend (what blue means, what each colour means) and a "How to use" note on any sheet other people will fill in.
- Protect formula cells (Review > Protect Sheet) if others will type in the sheet; leave input cells unlocked.

## 10. Shortcuts worth learning
| Keys | Action |
|---|---|
| Ctrl+T | Make a table |
| Ctrl+Shift+L | Turn filters on or off |
| Alt+= | AutoSum |
| F4 | Cycle absolute and relative references |
| Ctrl+; and Ctrl+Shift+; | Insert today's date and the time |
| Ctrl+` | Show formulas |
| Ctrl+Arrow, Ctrl+Shift+Arrow | Jump to or select to the edge of data |
| Ctrl+D, Ctrl+R | Fill down, fill right |
| Ctrl+E | Flash Fill |
| Alt, N, V | Insert PivotTable |
| F5, then Special | Go To Special |
| Ctrl+Shift+Enter | Enter a legacy array formula |

## 11. Power Query, Power Pivot and macros (when formulas stop being enough)
- **Power Query** (Data > Get Data): import from files, folders or the web; trim, split, unpivot, merge and append; the steps are saved and re-run on Refresh. Best for monthly files with the same layout.
- **Power Pivot / Data Model:** relate several tables and write measures in DAX (`SUM`, `CALCULATE`, `DISTINCTCOUNT`). Best when one pivot needs data from several tables.
- **Macros:** record a repeated task, then clean up the code. Save as .xlsm. Prefer formulas when they can do the job; macros do not run in Excel for the web or Google Sheets.
- **Real slicers and timelines** work on PivotTables and Tables (Insert > Slicer).

## 12. A note on identity numbers (CNIC)
Pakistani CNICs are normally written `12345-1234567-1` (13 digits). Older cards use an 11-digit number such as `333-15-034961`. In the practice data, both
styles appear, with dots, dashes and plain digits, so clean before comparing. The first digit is commonly described as the province and the first five
digits as a locality code; treat that as background only, and do not rely on it. The workbook region formula instead learns from earlier records. Personal
ID numbers are sensitive: keep files private, show only what is needed, and do not paste real ones into public tools.
