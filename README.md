# Fair Usage V2 – Google Apps Script Toolkit

This repository contains the Google Apps Script code that powers the **Tech Fee Tools** custom menu inside a Fair Usage tracking Google Sheet. The scripts join client revenue with tech-fee data, generate a fair-usage allocation table, and keep the required configuration tabs in sync. A companion helper (`OnCrawlBuget.gs`) populates OnCrawl crawl budgets from an exports sheet or lets you override budgets manually.

## Contents
- `Code.gs` – main menu and fair-usage logic (AccuRanker, Semrush, OnCrawl allocation, Setup tab bootstrapper).
- `OnCrawlBuget.gs` – helper for populating the `OnCrawl Monthly URL Budget` column on the `Adjustments` sheet.

## Spreadsheet Requirements
The scripts assume a Google Sheet with the following tabs:
- `SEO Client Revenue` – revenue per account/market (headers on row 1, data starting row 2). Columns C–E map to 2024–2026 revenue.
- `Tool Revenue` – tech-fee revenue by account (headers on row 4, data starting row 5). Columns B–D map to 2024–2026 revenue.
- `Setup` – configuration table (created automatically if missing).
- Optional: `Adjustments` and `OnCrawl Stats` – required if you run the OnCrawl budget helper.

## Installation
1. Open the target Google Sheet and launch **Extensions → Apps Script**.
2. Replace the default `Code.gs` with the contents of this repository’s `Code.gs`. Add a second file for `OnCrawlBuget.gs`.
3. Save your project and reload the spreadsheet to expose the **Tech Fee Tools** menu.

## Custom Menu
On spreadsheet open (`onOpen`), a **Tech Fee Tools** menu is added with three actions:

- **Build Tech Fee Join** (`Build_Tech_Fee_Join`)  
  Prompts for a year (2024–2026), then joins the `SEO Client Revenue` and `Tool Revenue` tabs to create `Revenue vs Tech Fee – {year}`. The output flags whether each account is paying a tech fee and formats the table with filters and number styles.

- **Refresh Fair-Usage Table** (`Build_FairUsage_ForYear`)  
  Rebuilds the `Tech Fair-Usage – {year}` sheet. It:
  - Loads or creates the `Setup` tab and reads capacity/percentage settings.
  - Calculates revenue tiers, region bands, and site-size multipliers per account.
  - Splits AccuRanker capacity into base and contributor pools, respecting tier ceilings and a global buffer.
  - Derives Semrush keyword caps, OnCrawl cadence, and starter crawl budgets.
  - Outputs a fully formatted summary with frozen headers and filters.

- **Create/Update Setup Tab** (`EnsureSetupTab_`)  
  Creates the `Setup` configuration tab (and pads it to four columns) if it is missing or sparsely populated. It also back-fills newer sections—crawl cadence rules, OnCrawl starter caps, and site size multipliers—on older sheets.

## Setup Tab Structure
`EnsureSetupTab_` expects (and seeds) the following blocks, each separated by a blank row:

| Block | Notes |
| --- | --- |
| `ACCURANKER_CAPACITY`, `SPLIT_BASE_PCT`, `SPLIT_POOL_PCT`, `SPLIT_BUFFER_PCT` | Control the total AccuRanker slots and Base/Contributor/Buffer split. |
| `Revenue Tiers` | Defines tier min/max revenue, plus base & ceiling allocations. |
| `Regional Bands` and `Market+Regional Band` | Associate markets with multipliers that influence contributor pool share. |
| `Semrush Caps` | Keyword caps per tier for paying vs non-paying accounts. |
| `Crawl Cadence Rules` | Default OnCrawl cadence strings. |
| `OnCrawl Starter Caps` | Base crawl budget defaults per tier (paying and non-paying). |
| `Site Size Multipliers` | Optional multipliers tied to site size buckets. |
| `Account+Site Size` | Overrides to map specific accounts to size buckets. |

Values can be edited directly in the sheet; the scripts read display values so formulas are supported.

## OnCrawl Budget Helper
`populateOncrawlMonthlyBudget` (in `OnCrawlBuget.gs`) fills the `OnCrawl Monthly URL Budget` and `Budget Source` columns on the `Adjustments` sheet.

Requirements:
- `Adjustments` sheet with headers: `Domain`, `Override OnCrawl Budget`, `OnCrawl Monthly URL Budget`, `Budget Source`.
- `OnCrawl Stats` sheet with headers: `Domain`, `Monthly URL Budget`, `Avg Daily URLs`, `Crawl Days In Month`.

Logic:
1. Preserve any manual overrides in `Override OnCrawl Budget`.
2. Otherwise, pull matching OnCrawl stats.
3. If the monthly budget is blank but daily averages exist, multiply by `Crawl Days In Month` (default 30) to derive a monthly value.
4. Record where each value came from (`override`, `oncrawl`, or `missing`).

The helper can be called manually from Apps Script or wired to a custom menu/trigger if desired.

## Development Notes
- Helper utilities (`safeStr_`, `toNumber_`, `getLastRow_`, etc.) keep parsing robust when spreadsheets contain blanks or formatted numbers.
- The allocation algorithm enforces the AccuRanker capacity cap by scaling contributor allocations first, then base allocations, to stay under the buffer-adjusted ceiling.
- All outputs rewrite their target sheets entirely, apply number formats, freeze headers, and re-create filters for easy analysis.

## Next Steps
- Review and tailor the seeded `Setup` values to match current commercial policy.
- Add simple buttons or triggers in the spreadsheet UI for the OnCrawl helper if the workflow needs it regularly.
- Consider protecting the `Setup` sheet once configured to avoid accidental edits.
