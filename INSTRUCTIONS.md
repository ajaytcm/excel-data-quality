# Step-by-Step — Build the Cleaned Workbook

Time: about 45–60 minutes. Tool: Microsoft Excel (desktop or web). Work on a copy — the Raw_Data sheet is never edited.

## Step 1 — Load the raw data
1. Open `data/messy_supplier_orders.csv` in Excel.
2. Save the file as `cleaned_workbook.xlsx` in this folder.
3. Rename the sheet to **Raw_Data**. This sheet stays untouched from now on.

## Step 2 — Start the clean copy
1. Right-click the Raw_Data tab → Move or Copy → tick **Create a copy** → rename the copy **Clean_Data**.
2. Select all data on Clean_Data, then **Data → Remove Duplicates** (all columns ticked). Excel should report 1 duplicate removed (the exact PO-2003 copy). Note the count — it goes in your findings.

## Step 3 — Standardise the text columns
1. **Supplier:** click an empty column, use `=TRIM(A2)` style formulas (adjust to the supplier column) to strip extra spaces, then Copy → Paste Special → Values over the original column. "Acme Industrial " and "Acme Industrial" must become one supplier.
2. **Status:** use Find & Replace on the status column: `received` → `Received`, `RECEIVED` → `Received`, `Recieved` → `Received`. Final allowed list: Received, In Transit, Pending, Cancelled.

## Step 4 — Fix the dates
1. The Order Date column mixes `2026-09-01`, `09/02/2026` and `2026/09/04`. Select the column → **Data → Text to Columns → Date (YMD)** where needed, then format the whole column as `yyyy-mm-dd`.
2. One row has the impossible date `32/13/2026` (PO-2012). You cannot fix a date you don't know — highlight the cell, and record it in the Validation sheet as an open issue to send back to the source. Do not invent a date.

## Step 5 — Build the Validation sheet
Create a sheet **Validation** that lists every rule, how you checked it, and how many rows fail:

| Rule | How to check (on Clean_Data) | Expected finding |
|------|------------------------------|------------------|
| Missing PO number | Filter PO column for blanks | 1 row (Delta Fasteners, 2026-09-05) |
| Quantity <= 0 | Filter or `=COUNTIF(qty_range,"<=0")` | 2 rows: −25 (PO-2007) and 0 (PO-2016, Cancelled — confirm this is the cancellation convention) |
| Missing unit cost | `=COUNTBLANK(cost_range)` | 1 row (PO-2009) |
| Price outlier | Add a helper column: usual cost per part via `=VLOOKUP(part, price_list, 2, FALSE)` against a small price list you type from the most common values; flag rows where actual cost <> usual cost | 1 row: Steel Plate at 380.00 vs 38.00 (PO-2008) — a decimal slip, correct it to 38.00 and note the correction |
| Near-duplicate PO | Sort by PO number; PO-2015 appears twice with different dates | Keep both flagged — different dates may be a split delivery. Record as "verify with supplier", do NOT delete |

## Step 6 — Correct what is safely correctable
On Clean_Data only: fix the 380.00 → 38.00 (note it in Validation), leave the unknown date and the missing PO/cost untouched and flagged. A data-quality rule worth remembering: **correct typos, flag unknowns — never invent values.**

## Step 7 — Summary PivotTable
1. Select Clean_Data → **Insert → PivotTable** on a new sheet named **Summary**.
2. Rows: Supplier · Values: Sum of line total (add a helper column `=qty*unit_cost` named Line Total first) and Count of PO Number.
3. Second PivotTable: Rows: Status · Values: Count of PO Number.
4. Check: total order count in the pivot = rows on Clean_Data. If they differ, a filter or a blank is hiding rows — find out why before finishing.

## Step 8 — Finish and upload
1. Add a text box or first-row note on Summary: data source, date cleaned, rules checked, open issues (missing PO, missing cost, impossible date, PO-2015 to verify).
2. Save, then upload `cleaned_workbook.xlsx` to this repo next to the README.

## How to talk about this project in an interview
- "I separate raw from clean and never edit raw — every correction is traceable."
- "I corrected typos and a decimal-slip price, but flagged the missing PO, missing cost and impossible date instead of inventing values."
- "The VLOOKUP price check is the same idea as a three-way match: compare what was charged against the expected price and flag the variance."
