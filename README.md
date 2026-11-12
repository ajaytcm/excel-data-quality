# Excel Data Quality — Supplier Orders Cleanup

A hands-on Excel project: take a messy supplier-orders extract, clean it, validate it, and summarise it — the same work as checking a vendor/PO file before it goes into an ERP or a report.

**Skills shown:** data cleaning, duplicate removal, text standardisation, formula-based validation (flags, VLOOKUP), PivotTable summaries.

## Files

- `data/messy_supplier_orders.csv` — the raw extract (20 rows, synthetic sample data) with planted defects: exact duplicates, a near-duplicate, a missing PO number, mixed date formats, one impossible date, inconsistent status spelling/casing, a negative quantity, a missing unit cost, a 10x price outlier, and a supplier name with a trailing space.
- `INSTRUCTIONS.md` — the step-by-step build. Follow it to produce the cleaned workbook.
- `cleaned_workbook.xlsx` — the finished workbook (added after completing the steps in INSTRUCTIONS.md).

## Workbook structure (target)

1. **Raw_Data** — the CSV exactly as received. Never edited.
2. **Clean_Data** — cleaned copy: duplicates removed, dates in one format, statuses standardised, supplier names trimmed.
3. **Validation** — formula flags for each rule (missing PO, qty <= 0, missing cost, price outlier vs. the part's usual cost) with a count of open issues.
4. **Summary** — PivotTable: spend by supplier and order count by status, from the clean data only.

## Defects planted in the raw file (find them with the steps, don't just read this list)

Duplicate PO-2003 · PO-2015 repeated with a different date · one row with no PO number · dates in three different formats plus one impossible date (32/13/2026) · status values `received` / `RECEIVED` / `Recieved` · quantity −25 · one blank unit cost · Steel Plate at 380.00 vs. its usual 38.00 (decimal slip) · "Acme Industrial " with a trailing space (breaks grouping until trimmed).
