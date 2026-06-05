# XLForge Web ⚒️🌐

**The [XLForge](https://github.com/kristic8998/xlforge) Excel automation toolkit, rebuilt for Excel on the web.**

Same one-call data cleaning and instant reports — no desktop Excel, no installs. Two flavors included:

| | Requires | Best for |
|---|---|---|
| **Office Scripts** (`scripts/`) | Automate tab (Microsoft 365 work/school) | Full automation: clean + report in one click |
| **LAMBDA formulas** (`formulas/`) | Nothing — any Excel on the web | Formula-native cleaning, zero permissions |

![TypeScript](https://img.shields.io/badge/Office%20Scripts-TypeScript-3178C6?logo=typescript&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-on%20the%20web-217346?logo=microsoftexcel&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-blue)

## Quick start (Office Scripts)

1. Open any workbook at [office.com](https://office.com) → **Automate** tab → **New Script**
2. Delete the starter code, paste in [`scripts/cleanData.ts`](scripts/cleanData.ts)
3. Rename the script to "XLForge Clean Data" → **Save script**
4. Open the sheet with messy data → **Run**

Repeat with [`scripts/buildReport.ts`](scripts/buildReport.ts) for the report builder (set `GROUP_COLUMN` / `VALUE_COLUMN` at the top to match your data).

> No Automate tab? Your Microsoft 365 plan doesn't include Office Scripts — use the [LAMBDA toolkit](formulas/LAMBDA-toolkit.md) instead, which works everywhere.

## What the cleaner does

One run on the active sheet:

- Trims whitespace, collapses double spaces, strips non-printing characters and non-breaking spaces (the web copy-paste classics)
- Auto-detects columns of text dates (`2026-01-15`, `1/15/2026`, `Jan 15 2026`) and converts them to real Excel dates
- Applies Proper Case to text
- Removes exact duplicate rows, keeping the first
- Writes a timestamped audit trail to a `CleanLog` sheet — you always know what changed

## What the report builder does

- Groups any column by any value column (sum / count / average)
- Outputs a styled, banded table with a live `=SUM()` total row
- Adds a column chart automatically
- Never touches your source data — everything lands on a new sheet

## Test it

[`test-data/messy-sales.csv`](test-data/messy-sales.csv) contains 200 rows of synthetic sales data with every problem the cleaner handles. Upload it to OneDrive, open with Excel on the web, and run the scripts.

## Design notes

- **Batched I/O**: one `getValues()` read and one `setValues()` write per run — Office Scripts round-trips to the cloud, so per-cell access is the #1 performance killer
- **Dates as serials**: text dates convert to true Excel serial numbers (computed in UTC to avoid timezone drift), then get a `yyyy-mm-dd` number format
- **Audit-first**: like the VBA original, every run logs what it changed — automation you can't audit is automation you can't trust
- **No hardcoded totals**: report totals are live formulas, so the sheet stays honest if someone edits a number

## Relationship to XLForge (VBA)

This is a sibling of [kristic8998/xlforge](https://github.com/kristic8998/xlforge) — same philosophy (array processing, audit logging, defensive defaults), different runtime. The VBA version targets desktop Excel; this one targets the browser.

## License

MIT — see [LICENSE](LICENSE).
