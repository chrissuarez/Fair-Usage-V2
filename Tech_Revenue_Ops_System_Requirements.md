# **Tech Revenue Ops System: Technical Requirements**

**Objective:** Build a Google Apps Script application to centralize Earned Media revenue tracking, automate "Fair Usage" tier enforcement, and identify commercial risks (e.g., high-revenue clients paying $0 tech fees).

Developer Profile: Senior Dev using VS Code \+ AI Assistant.  
Target Environment: Google Sheets (Backend/UI) \+ Apps Script (Logic).

## 🔑 Shared Keys & Config Tabs (applies to all phases)

* **Primary Key:** `Opportunity_UID` (Salesforce Opportunity Id; if missing, generate a hash of Account + Opportunity Name). Every sheet must carry this column.  
* **Config Tabs (read-only):**  
  * `Config_Currency`: Currency, Rate_to_USD (static), Last_Updated (manual).  
  * `Config_Tech_SKU_Pricing`: Tier (A/B/C/D), SKU_Name (Tech Starter/Tech Pro), Monthly_USD, Annual_USD, Notes.  
  * `Config_Params`: Today(), Renewal_Lookahead_Days (default 90), Partial_Month_Mode (`PRORATE` or `SIMPLE`).  
* **Audit:** `Import_Log` sheet with columns Source (Estimate/TechFee), Imported_By, Timestamp, Row_Count, File_Name, Validation_Status, Errors.

## **🏗 Phase 1: Data Architecture & Ingestion (Branch: feature/data-pipeline)**

**Goal:** Establish a robust "Shadow Table" pattern. We must ingest raw Salesforce data but allow manual overrides without breaking the pipeline on the next export.

### **1.1 Data Source Definitions**

We require two primary CSV ingestions.

* **Source A: Estimate/Resource Export (Revenue Source)**  
  * *Purpose:* Calculates true SEO/PR revenue based on staff hours, ignoring messy Opportunity Product lines.  
  * *Key Columns:* Account, Opportunity Name, Resource Role, Resource Region, Hours, Bill Rate, Start Date, End Date.  
* **Source B: Tech Fee Product Export (Contribution Source)**  
  * *Purpose:* Identifies who is actually paying for tools.  
  * *Key Columns:* Account, Opportunity Name, Product Name (Look for "Tech", "Tools"), Amount, Start, End.

### **1.2 The "Shadow Table" Architecture**

Create a script to manage three distinct sheet types for *each* data source.

1. **RAW\_Data\_Import:**  
   * *Behavior:* Locked. Deleted and replaced 100% on every new CSV import.  
   * *Action:* Script parses CSV and pastes values here.  
   * *Required Columns:* Opportunity\_UID, Account, Opportunity Name, Resource Role, Resource Region, Hours, Bill Rate, Start Date, End Date, Currency, Tech Fee Product?, Product Amount.  
2. **MANUAL\_Overrides:**  
   * *Behavior:* Persistent User Input. Data validation on Field to Override (Revenue, Tech Fee Status, Start Date, End Date, Currency, Product Amount).  
   * *Columns:* Opportunity\_UID (Key), Field to Override, New Value, Reason, Entered By, Timestamp.  
   * *Logic:* User adds a row here if Salesforce is wrong (e.g., "Deckers actually pays tech via a generic retainer line").  
3. **MASTER\_Ledger (The Computed View):**  
   * *Behavior:* Generated via Script/Formula.  
   * *Logic:* Value \= IF(Override\_Exists, Override\_Value, Raw\_Value), with currency normalized to USD via `Config_Currency`.  
   * *Output:* This sheet is the single source of truth for all dashboards.

### **1.3 CSV Import Workflow**

1. Validate headers against required columns; fail fast and log to `Import_Log` if missing.  
2. Clear and refill RAW\_Data\_Import with values-only, preserving header row formatting.  
3. Stamp `Import_Log` with row count and validation result.  
4. Recompute MASTER\_Ledger by applying overrides and currency conversion.  
5. Surface any records missing Opportunity\_UID or Start/End dates in a "Data Quality" view.

## **🧠 Phase 2: The "Earned Media" Revenue Engine (Branch: feature/revenue-logic)**

**Goal:** Deduce specific capability revenue (SEO vs. PR vs. ASO) by analyzing resource roles.

### **2.1 The Categorization Logic**

Create a helper function categorizeRevenue(roleString):

* **Input:** "SEO \- Manager", "Public Relations \- Director", "ASO \- Analyst", "Content \- Lead".  
* **Logic:**  
  * IF role contains "SEO" → Tag as **SEO Revenue**.  
  * IF role contains "Public Relations" OR "PR" → Tag as **Digital PR Revenue**.  
  * IF role contains "ASO" OR "App Store" → Tag as **ASO Revenue**.  
  * *Fallback:* Tag as "Other/Shared".
* **Implementation Sketch:**  
  ```javascript
  function categorizeRevenue(roleString) {
    const role = (roleString || '').toUpperCase();
    if (role.includes('SEO')) return 'SEO Revenue';
    if (role.includes('PUBLIC RELATIONS') || role.includes('PR')) return 'Digital PR Revenue';
    if (role.includes('ASO') || role.includes('APP STORE')) return 'ASO Revenue';
    return 'Other/Shared';
  }
  ```

### **2.2 The Run-Rate Calculator**

Create a function calculateMonthlyRevenue():

* **Input:** Total Revenue (Hours \* Rate), Start Date, End Date.  
* **Logic:**  
  * Calculate Months\_Duration.  
  * Monthly\_Revenue \= Total\_Revenue / Months\_Duration.  
  * Partial month mode: if `Config_Params.Partial_Month_Mode` = `PRORATE`, compute day-level prorate for first/last month; else use simple division.  
  * Convert to USD using `Config_Currency` before calculations.  
* **Output:** Populate the MASTER\_Ledger with a month-by-month breakdown (Jan '25, Feb '25...) for accurate forecasting.

* **Implementation Sketch:**  
  ```javascript
  function calculateMonthlyRevenue(amount, startDate, endDate, partialMode = 'SIMPLE') {
    const start = new Date(startDate), end = new Date(endDate);
    const months = Utilities.monthsBetween(start, end) || 1; // implement helper
    if (partialMode === 'PRORATE') {
      return spreadProratedByDay(amount, start, end);
    }
    return spreadEvenly(amount, months);
  }
  ```

## **⚖️ Phase 3: The "Entitlement vs. Reality" Engine (Branch: feature/health-check)**

**Goal:** Compare the *Calculated Revenue* (Phase 2\) against the *Tech Policy* to flag risks.

### **3.1 Tier Assignment Logic**

For every Account in MASTER\_Ledger, assign a **"Target Tier"** based on *Annualized SEO Revenue*.

* **Tier A:** \> $500k  
* **Tier B:** $200k \- $499k  
* **Tier C:** $75k \- $199k  
* **Tier D:** \< $75k

*Target fee mapping (USD, annualized):*  
* Tier A → Tech Pro @ $12k  
* Tier B → Tech Pro @ $12k  
* Tier C → Tech Starter @ $6k  
* Tier D → Tech Starter @ $2.4k

### **3.2 The "Health Check" Algorithm**

Create a function generateClientActionPlan(client) that outputs a recommendation string.

**Logic Blocks:**

* **Scenario: The "High Value Freeloader"**  
  * *Condition:* Revenue \> $200k (Tier A/B) AND Tech\_Fee\_Paying \== False.  
  * *Status:* 🔴 **CRITICAL MISS**  
  * *Plan:* "Client is consuming Enterprise resources. Must add 'Tech Pro' ($12k) at next renewal. Risk of service degradation if not actioned."  
* **Scenario: The "Under-Payer"**  
  * *Condition:* Target\_Fee\_SKU ($6k) \> Actual\_Fee ($1.5k).  
  * *Status:* 🟠 **GAP IDENTIFIED**  
  * *Plan:* "Legacy Pricing detected. Propose 'Glide Path': Renewal Year 1 at $3k (50% discount), Year 2 at full price."  
* **Scenario: The "Small Client / High Tax"**  
  * *Condition:* Revenue \< $50k AND Target\_Fee ($6k) \> 10% of Revenue.  
  * *Status:* ⚠️ **FEE TOO HIGH**  
  * *Plan:* "Standard fee exceeds 10% of retainer. Downgrade target to 'Tech Starter' ($2.4k) to preserve deal margin."  
* **Scenario: The "Compliant"**  
  * *Condition:* Actual\_Fee \>= Target\_Fee.  
  * *Status:* 🟢 **HEALTHY**  
  * *Plan:* "No action needed. Client fully funds their tier."

* **Outputs:** Health status, Recommended Action, Target\_Fee, Actual\_Fee, Fee Variance, Tier. Persist these in MASTER\_Ledger for UI use.

## **🖥 Phase 4: UI & Visualization (Branch: feature/dashboard-ui)**

**Goal:** A "Control Room" dashboard in Sheets to visualize the data from Phase 3\.

### **4.1 The "Portfolio Health" Dashboard**

Create a visual sheet showing:

1. **Client Name**  
2. **Annual SEO Revenue** (Calculated)  
3. **Target Tech Tier** (A/B/C/D)  
4. **Target Tech Fee** ($)  
5. **Actual Tech Fee** ($)  
6. **Variance** (Target \- Actual)  
7. **Status** (🔴/🟠/🟢/⚠️)  
8. **Recommended Action** (Text string from 3.2)

### **4.2 The "Renewal Radar"**

* Filter the Master Ledger by End Date.  
* Highlight clients renewing in the next **90 Days** who have a 🔴 or 🟠 status.  
* *User Story:* "As a Head of Ops, I want to see a list of 'At Risk' renewals so I can email the Client Lead today."

## **🛠 Technical Notes for Developer**

* **Currency Normalization:** Ensure all revenue inputs (GBP, EUR, USD) are converted to a single base currency (USD) for Tier calculations. Use a static exchange rate table in a Config sheet to avoid API complexity.  
* **Code Structure:**  
  * Ingestion.gs (CSV Parsing)  
  * RevenueLogic.gs (Role parsing & Monthly spread)  
  * PolicyEngine.gs (Tier logic & Action Plans)  
  * UI.gs (Dashboard generation)  
* **Error Handling:** If an Opportunity has "SEO" hours but no "Tech Fee" product, default Tech\_Fee\_Paying to False but flag it as "Review SOW" in the Override sheet.

## ✅ MVP Exit Criteria

* Successful imports for both CSV types with non-destructive override handling and Import\_Log entries.  
* MASTER\_Ledger shows currency-normalized monthly revenue per opportunity and applies overrides.  
* categorizeRevenue and calculateMonthlyRevenue tested with sample roles/dates in Apps Script logs.  
* Tier assignment and generateClientActionPlan outputs present for every Account with visible status color.  
* Portfolio Health and Renewal Radar tabs render without manual formulas after a full pipeline run.
