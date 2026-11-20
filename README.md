# GB reconciliation: PGDRCVNT receipts vs container payments

Project plan v2, updated 20 November 2026. The live copy of this plan is the Google Doc
"PGDRCVNT Reconciliation: Project Plan v2"; this file mirrors it.

The Apps Script is [`PGDRCVNT_Simple.gs`](PGDRCVNT_Simple.gs). It is bound to the master Google Sheet
"GB ชน 2026" (Extensions > Apps Script). To update it, paste the file over the script there and save.

## Summary

The reconciliation runs on two tabs and one script. Each refno from the ERP receipts (PGDRCVNT) gets a
container (ตู้) typed once in the Refno tab. The Containers tab compares what was paid in the supplier
payment tabs with what was received. The only recurring manual work is choosing a ตู้ for new refnos,
and the sheet suggests the most likely ones. Since 20 November 2026 the import also runs by itself every
Monday morning and every run is logged.

## Where the data comes from

```
Formula SQL Server
   | reports/PGDRCVNT/daily_pgdrcvnt.py + upload_pgdrcvnt_claude.py   (repo emkrit/formula-erp-reports)
   v
Google Sheet PGDRCVNT, tab PGDRCVNT_Claude
   | IMPORTRANGE
   v
Google Sheet PGDRCVNT_Copy, tab Sheet1            <- the script reads this (never edits it)
   | PGDRCVNT_Simple.gs: Import & Clean (button, or every Monday 08:00)
   v
Master sheet: PGDRCVNT -> Refno -> Containers
   | every run
   v
Google Sheet "Claude project log update", tab Sheet1   (one row per run)
```

## Goal

- Import and clean the PGDRCVNT receipts with one click, or automatically every Monday.
- List every refno once, with its receive date, number of lines, qty and amount.
- Status per refno: Match when a ตู้ is entered, N/A for freebies and samples with no container, Pending otherwise.
- For Pending refnos, suggest the closest container by amount, only among containers whose ETD is within 3 weeks of the receive date.
- Compare paid and received per container.
- Never edit the payment tabs (CNS, HU, ZS, FX, NT), Summary or the source sheet. Data errors found there are reported to Em, who fixes them.

## Tabs

| Tab | Content | Who edits it |
|---|---|---|
| Refno | One row per refno: Supplier, refno, receive date, lines, qty, amount, ตู้, Status, qty and amount paid for that ตู้, diff qty, diff amount, three suggestions, notes | You type column G (ตู้) and column P (หมายเหตุ) only |
| Containers | One row per ตู้ from the payment tabs: ETD, qty, paid amount, received amount, diff, result, refnos | Script only |
| PGDRCVNT | Cleaned receipts, one row per receipt line. Column A = status, B = ตู้ of the line, C = price × qty | Script only |
| CNS, HU, ZS, FX, NT | Payment tabs, maintained by hand | You, as before |

The old tabs (Input, Helpers, Pending_Match, Summary_Auto, Container_Map, Total, CheckResult,
SKU_vs_Summary) are hidden, not deleted.

## Workflow

1. Make sure PGDRCVNT_Copy (tab Sheet1) shows the latest ERP receipts. It pulls them by IMPORTRANGE, so there is no file to export or upload.
2. Every Monday 08:00–09:00 (Bangkok) the script runs **Import & Clean** and **Refresh** by itself. You can also run **PGDRCVNT > 📥 Import & Clean** any time.
3. In Refno, the Pending rows are at the top. For each, type the ตู้ in column G (pick from the dropdown), `ZS4/ZS5` if the refno belongs to two containers, or `N/A` for a freebie or sample.
4. Run **PGDRCVNT > 🔄 Refresh** to update the container totals and the suggestions for the rows still Pending.
5. Check Containers for rows marked ต่าง (different) or ใกล้เคียง (close), and the log sheet for the Pending list.

Typed ตู้ are stored in Refno column G and kept on every import and refresh; you never type them twice.

## Weekly run and log

- Turn it on once: **PGDRCVNT > ⏰ ตั้งรันอัตโนมัติ (ทุกจันทร์ 8:00)**. Google asks for permission the first time.
  **⛔ ยกเลิกรันอัตโนมัติ** turns it off. Clicking the first item again replaces the schedule; it never creates a second one.
- Each run, scheduled or by button, appends one row to the log sheet "Claude project log update", tab Sheet1:

| Column | Value |
|---|---|
| B Finished (TH) | time of the run (dd/MM/yyyy HH:mm, Bangkok) |
| C Task | GB auto reconcile |
| D Trigger | `Script Auto run`, `Manual: Import & Clean` or `Manual: Refresh` |
| E Run days | Every Monday |
| F Status | Success or Error |
| G Message | lines imported, Match / N/A / Pending counts, container results, or the error |
| H Datebased | PGDRCVNT_Copy and its receive-date range |
| I Remark | `# Refno Pending = N: CNS 11244, FX 16060-1, …` |

- The scheduled run shows no pop-ups. If PGDRCVNT_Copy has fewer than half the rows of the last import
  (for example IMPORTRANGE has not loaded), it stops without changing anything and logs Error. Then check
  PGDRCVNT_Copy and click Import & Clean yourself (that one asks before going on).
- If the log sheet can't be opened, the run still finishes; only the log row is missing.

## Cleaning rules

1. **Read** PGDRCVNT_Copy and match its Thai column headers to the keys the script uses (the ERP document number is column `code`; `book` is always 0002).
2. **One value per cell**: the source is a sheet, so inch marks such as `(1")` and commas in product names are read correctly.
3. **Remove empty rows**: rows with no document code or no run.
4. **Remove duplicate lines.** A receipt line is identified by `ab` + document number + `run`. When the same key appears more than once, the copy with an invoice date is kept; otherwise the first copy. The original row order is then restored.
5. **Check each row.** A row is skipped when run is not a whole number, or prodprice or qty is not a number. Skipped rows are listed by no/run after import. If no row can be read, the import stops before anything is changed.
6. **Import suppliers only** (coorcodes below). If none are found, the import stops.
7. **Write real dates** (the master sheet is in US locale, so text dates would swap day and month). Refnos and codes are written as text, so 09739 keeps its leading zero.
8. Refnos received earlier than (the supplier's first ETD − 7 days) are older than the payment tabs and are left out of Refno, unless a ตู้ was typed. Refnos missing from newer data are kept with their last totals only when a ตู้ was typed.

## Matching rules

- **Supplier** from coorcode: CNS = จ.006, ค.022; FX = อ.015, อ.085; HU = อ.042; ZS = ฮ.005; NT = น.008.
- **Refno** is compared without leading zeros and case (06865 = 6865).
- **Line amount** = prodprice × qty. **Paid** = the lowercase `amount` column of the payment tab. **ETD** = the earliest ETD of the container.
- **Two containers** (`ZS4/ZS5`): the lines of the refno are split in passes, best evidence first: same product code and amount → same amount → same receipt document as lines already placed → same product code, closest amount → container with the closest ETD.
- **Suggestions** for a Pending refno: same supplier, ETD within 21 days of the receive date, ranked by how close the refno amount is to the container's amount not yet received. `พอดี` = exact. Amount 0 is suggested as N/A.
- **Container result**: ✅ ตรง within max(1, 0.1%); 🟡 ใกล้เคียง within 2%; ❌ ต่าง beyond; ⏳ ยังไม่รับ when no refno is matched yet.
- **Refno diff colours** (columns K, L): green within the ✅ limit, yellow within 2%, red beyond.

## Issues found and fixed

| Date | Issue | Fix |
|---|---|---|
| 3 Nov | Inch marks `(1""` in the CSV were read as an escaped quote; lines were lost (8433 runs 2–3, every ZS container short) | New reader; since 19 Nov the source is a sheet, so no CSV parsing |
| 3 Nov | Text dates swapped day and month (US locale) | Real dates are written |
| 3 Nov | Rows with commas in the product name were dropped | Row check by value, not column count |
| 3 Nov | A check rejected every row; import read 0 rows and Refresh wiped Refno | Import stops before any change; Refresh stops when there are no lines |
| 3 Nov | A PDF saved as PGDRCVNT.csv, and a local-supplier report, were imported | File-type and supplier checks |
| 3 Nov | ZS4/ZS5 split wrong (ZS row 32 has no product code) | Multi-pass split over all lines of the refno; row 32 reported to Em |
| 3 Nov | 2025 refnos in Refno | Removed; rule 8 above |
| 19 Nov | Source moved from CSV to Google Sheet PGDRCVNT_Copy | Script reads the sheet directly |
| 20 Nov | Manual weekly run | Monday 08:00 trigger + run log |

## Open items

1. Type the ตู้ for the Pending refnos that have an exact suggestion: 11244 → C25, 16060-1 → C23_YiWu, 17355 → H9, 03278 → N14.
2. C21 + Yiwu is 7,414 short (possibly refno 15051-1); F5_W.F2 is 96 over (refno 11768-2).
3. ZS tab row 32 has a blank product code (fix in the payment tab).
4. Optional: a button that fills every `พอดี` suggestion at once.
5. Delete the hidden old tabs once the new setup has been used for a full cycle.

## Settings at the top of the script

| Setting | Value | Meaning |
|---|---|---|
| SOURCE_SHEET_ID / SOURCE_TAB | PGDRCVNT_Copy / Sheet1 | Where receipts are read from |
| WINDOW_DAYS | 21 | Max days between ETD and receive date for a suggestion |
| BEFORE_FIRST_ETD_DAYS | 7 | Refnos received this long before a supplier's first ETD are left out |
| SUGGEST_N | 3 | Suggestions per Pending refno |
| OK_ABS, OK_PCT, NEAR_PCT | 1, 0.1%, 2% | Limits for ✅ and 🟡 |
| LOG_SHEET_ID / LOG_TAB | Claude project log update / Sheet1 | Where each run is logged |
| WEEKLY_HOUR, TZ | 8, Asia/Bangkok | Weekly run time (Google starts it within that hour) |

## Files

| File | What |
|---|---|
| `PGDRCVNT_Simple.gs` | The Apps Script bound to the master sheet |
| `tests/flow_test.js` | Checks the weekly run, the log row and the schedule buttons with stand-ins for Google services |

## Testing

The pure parts (`cleanRows`, `compute`, `summarize`, `logRow`) are exported for Node when `module` exists,
so they can be tested without Google services:

```
node tests/flow_test.js
```
