# Fair Usage V2 – Google Apps Script Toolkit

This repository contains the Google Apps Script code that powers the **Tech Fee Tools**, **Revenue Ops**, and **Cashflow Tools** inside the Fair Usage tracking Google Sheet. The scripts handle:
1.  **Fair Usage Allocations:** Deterministic assignment of tool limits (AccuRanker, Semrush, OnCrawl) based on client revenue tiers.
2.  **Revenue Ops Pipeline:** sophisticated ingestion of data streams (Estimates, Tech Fees, Opportunities, Projects) into a master ledger to track portfolio health and renewals.
3.  **Tech Cashflow Forecasting:** Projects monthly tech revenue and compares it against burn rates.

## Contents
- `Code.gs` – Menu definitions, Web App serving, and core Fair Usage logic.
- `TechRevenueOps.gs` – The financial engine. Handles data ingestion (Estimates, Opps), master ledger construction, and dashboard generation (Portfolio Health, Renewal Radar).
- `TechCashflow.gs` – Logic for projecting contract revenue onto a monthly timeline (`generateTechRunRate`).
- `TechFinanceTool.gs` – Bridging logic for the P&L dashboard and Tool Transaction ledger.
- `EmailImport.gs` – Utilities to fetch CSV reports directly from Gmail labels.

## Spreadsheet Requirements
The scripts assume a Google Sheet with the following key tabs (many are auto-created if missing):
- **Source Data (Manual/Imported):**
    - `SEO Client Revenue` & `Tool Revenue`: Legacy tabs for the basic Fair Usage check.
    - `Estimate_RAW_Data_Import`: **[Manual Entry]** Raw estimate data.
    - `Estimate_MANUAL_Overrides`: User-defined revenue overrides.
    - `Projects_RAW_Data_Import`: Project details imported from external sources.
    - `Opps_and_CRs_RAW_Import`: Opportunity hierarchy imported from email.
- **Configuration:**
    - `Setup`: Fair usage configuration (Tiers, Caps, Multipliers).
    - `Config_Currency`, `Config_Tech_SKU_Pricing`, `Config_Params`: Revenue Ops settings.
- **Outputs:**
    - `MASTER_Ledger`: The unified source of truth for all revenue.
    - `Tech Fair-Usage – {year}`: The final allocation table for stakeholders.
    - `Tech Cashflow Forecast 2025-26`: Monthly revenue projections.

## Custom Menu
On spreadsheet open (`onOpen`), two custom menus are added:

### **1. Fair Usage Tools**
Core utilities for managing allocations and the revenue pipeline.
- **Build Tech Fee Join** (`Build_Tech_Fee_Join`)  
  Joins `SEO Client Revenue` and `Tool Revenue` to audit which clients are paying tech fees.
- **Refresh Fair-Usage Table** (`Build_FairUsage_ForYear`)  
  Rebuilds the allocation table (`Tech Fair-Usage`), applying tier logic, site size multipliers, and crawler bonuses from the `Setup` tab.
- **Create/Update Setup Tab** (`EnsureSetupTab_`)  
  Regenerates the `Setup` tab if missing.
- **Create Revenue Ops Shadow Tables**  
  Ensures internal "Shadow" tables exist for data ingestion.
- **Refresh Revenue Ops Pipeline** (`refreshRevenueOpsPipeline`)  
  **Main Action:** Imports data (Opps/Projects), rebuilds the `MASTER_Ledger`, and updates `Portfolio Health` / `Renewal Radar` dashboards.
- **Import JF Buy Data**  
  Imports purchasing data from the external "JF Buy" sheet.

### **2. Tech Finance Admin**
Advanced financial dashboards and cashflow tools.
- **Run Full Update (Step 1 & 2)**  
  Runs both the data engine generation and dashboard build in sequence.
- **Step 1: Generate Data Engine** (`generateTechRunRate`)  
  Forecasts monthly revenue 2025-26 based on contract dates.
- **Step 2: Build Dashboard** (`buildDashboard`)  
  Updates the "Upgrade Predictor" visual dashboard.
- **Generate Tool Cost Timeline**  
  Projects tool costs monthly based on transaction logs.
- **Generate P&L Dashboard**  
  Builds the Profit & Loss view.

## Web App Interface
The project includes a comprehensive Web App (`index.html`) accessible via the script deployment. It provides a UI for:
- **Fair Usage Generation:** Triggering the allocation build.
- **Revenue Ops Control:** Viewing portfolio health and managing revenue overrides.
- **Tech Entitlements:** Viewing client entitlements based on the 2025 rate card.
- **Cashflow Dashboard:** Visualizing revenue vs. burn rate with scenario planning.
- **Settings:** Managing tool configurations and tier definitions.

## Setup Tab Structure
The `Setup` tab drives the Fair Usage logic. Key configuration blocks include:
- **Client Tier Matrix:** Revenue thresholds for Tiers A–D.
- **AccuRanker Caps:** Fixed keyword limits for paying vs. non-paying clients per tier.
- **Semrush Caps:** Strict keyword limits.
- **OnCrawl Starter Caps:** URL budgets (monthly) for new crawls.
- **Website Size Multipliers:** Adjusts allocations based on "Small", "Medium", "Large" site designations.

## Account Config
The system uses an **Account Config** tab (seeded by `EnsureAccountConfigTab_`) to manage high-level overrides:
- `Active?`: Force-stop allocations for churned clients.
- `Site Size`: Override the default size multiplier.
- `Own Crawler?`: Grants AccuRanker bonuses and skips OnCrawl allocations.
- `OneSearch Account?`: Adds specific keyword bonuses.

## Development Notes
- **Shadow Tables:** The Revenue Ops system uses a "Shadow Table" pattern (`RAW` -> `Overrides` -> `MASTER`) to ensure data integrity while allowing manual corrections.
- **Waterfall Dates:** `TechRevenueOps.gs` implements complex logic to handle Change Request (CR) dates vs. Parent Opportunity dates to prevent double-counting revenue.
