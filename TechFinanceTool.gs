/**
 * Tech Finance Tool - Master Controller
 * Contains all logic to generate the Data Engine and Dashboard.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tech Finance Admin')
    .addItem('Run Full Update (Step 1 & 2)', 'runFullUpdate')
    .addSeparator()
    .addItem('Step 1: Generate Data Engine', 'generateTechRunRate')
    .addItem('Step 2: Build Dashboard', 'buildDashboard')
    .addItem('Generate Tool Cost Timeline', 'buildToolCostTimeline')
    .addSeparator()
    .addItem('Create Tool Transactions Tab', 'ensureToolTransactionsTab')
    .addToUi();
}

function runFullUpdate() {
  try {
    generateTechRunRate(); // Step 1
    buildDashboard();      // Step 2
    SpreadsheetApp.getUi().alert("✅ Success! Both 'Tech Cashflow Forecast' and 'Upgrade Predictor' have been updated.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.message);
  }
}

// ==========================================
// STEP 1: DATA ENGINE (TechCashflow.gs)
// ==========================================
function generateTechRunRate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // CONFIG: EXACT SHEET NAME MATCHING
  const SOURCE_SHEET_NAME = "SEO Revenue from Closed Won Opps - Tech Fees Per Client"; 
  const TARGET_SHEET_NAME = "Tech Cashflow Forecast 2025-26";
  
  const START_PROJECTION_DATE = new Date("2025-01-01");
  const END_PROJECTION_DATE = new Date("2026-12-31");
  
  // 1. DATA VALIDATION
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error(`Could not find source sheet: '${SOURCE_SHEET_NAME}'. Please ensure your CSV is imported and renamed exactly.`);
  }
  
  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row
  
  // 2. COLUMN MAPPING (ADJUSTED FOR ACCURACY)
  // Headers based on your CSV structure:
  // 0:Account, 2:Start, 3:End, 15:Tech Fee (Final)
  const COL_ACCOUNT = 0;       
  const COL_START = 2;         
  const COL_END = 3;           
  const COL_TECH_FEE = 15; // CRITICAL: Index 15 is 'Tech Fee (Final)', NOT 'SEO Revenue'
  const COL_MARKET = 9;        
  
  // 3. GENERATE TIMELINE HEADERS
  const monthHeaders = getMonthHeaders(START_PROJECTION_DATE, END_PROJECTION_DATE);
  let outputRows = [];
  
  // 4. PROCESS CLIENTS
  data.forEach(row => {
    let clientName = row[COL_ACCOUNT];
    let startDate = new Date(row[COL_START]);
    let endDate = new Date(row[COL_END]);
    let totalFee = parseCurrency(row[COL_TECH_FEE]); // Uses helper to clean string "$1,000" -> 1000
    let market = row[COL_MARKET];
    
    // Validation: Must have fee and valid dates
    if (!totalFee || totalFee === 0 || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return;
    }
    
    // Calculate Monthly Run Rate
    let monthsDuration = monthDiff(startDate, endDate);
    if (monthsDuration < 1) monthsDuration = 1;
    let monthlyRevenue = totalFee / monthsDuration; 
    
    let clientRow = [clientName, market, startDate, endDate, totalFee, monthlyRevenue];
    
    // Spread Revenue Across Months
    monthHeaders.dates.forEach(monthDate => {
      // Logic: Client is active if the projection month overlaps with contract start/end
      let monthEnd = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0);
      
      if (monthDate <= endDate && monthEnd >= startDate) {
        clientRow.push(monthlyRevenue);
      } else {
        clientRow.push(0);
      }
    });
    
    outputRows.push(clientRow);
  });
  
  // 5. WRITE OUTPUT
  let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (targetSheet) targetSheet.clear();
  else targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
  
  let finalHeaders = ["Client", "Market", "Start Date", "End Date", "Total Fee", "Monthly Revenue", ...monthHeaders.labels];
  
  // Write Headers
  targetSheet.getRange(1, 1, 1, finalHeaders.length)
    .setValues([finalHeaders])
    .setFontWeight("bold")
    .setBackground("#EFEFEF");
  
  // Write Data
  if (outputRows.length > 0) {
    targetSheet.getRange(2, 1, outputRows.length, finalHeaders.length).setValues(outputRows);
    // Format Currency Columns (From Col 5 onwards)
    targetSheet.getRange(2, 5, outputRows.length, finalHeaders.length - 4).setNumberFormat("$#,##0.00");
  }
  
  targetSheet.setFrozenRows(1);
  targetSheet.setFrozenColumns(2);
}

// ==========================================
// STEP 2: DASHBOARD BUILDER (DashboardBuilder.gs)
// ==========================================
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DASHBOARD_NAME = "Upgrade Predictor";
  const DATA_SHEET_NAME = "Tech Cashflow Forecast 2025-26"; 
  const START_PROJECTION_DATE = new Date("2025-01-01");
  const END_PROJECTION_DATE = new Date("2026-12-31");
  
  // Verify Data Sheet Exists
  if (!ss.getSheetByName(DATA_SHEET_NAME)) {
    throw new Error("Data Sheet missing. Run Step 1 first.");
  }
  
  let sheet = ss.getSheetByName(DASHBOARD_NAME);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(DASHBOARD_NAME);
  
  // STYLING
  const titleStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(14).setForegroundColor("#1a73e8").build();
  const burnColor = "#fce8e6"; 
  const revenueColor = "#e6f4ea";
  const headerRowColor = "#f1f3f4";
  
  // --- SECTION A: CONFIG ---
  sheet.getRange("A1").setValue("1. Tech Burn Rate Config").setTextStyle(titleStyle);
  const costData = [
    ["Component", "2025 Monthly Cost", "2026 Monthly Cost", "Notes"],
    ["Base Stack (AccuRanker + Semrush)", 13144, 13144, "=(57732+100000)/12"],
    ["Upgrade (Semrush AIO - 2 Workspaces)", 0, 3000, "Upgrade lands in 2026"],
    ["Total Burn Rate", "=SUM(B3:B4)", "=SUM(C3:C4)", ""]
  ];
  sheet.getRange("A2:D5").setValues(costData);
  sheet.getRange("A2:D2").setFontWeight("bold").setBackground(headerRowColor);
  sheet.getRange("B3:C5").setNumberFormat("$#,##0");
  sheet.getRange("B5:C5").setBackground(burnColor).setFontWeight("bold");

  // --- SECTION B: PIPELINE ---
  sheet.getRange("F1").setValue("2. Tech Pipeline (Predicted Revenue)").setTextStyle(titleStyle);
  const pipelineData = [
    ["Client Target", "Est. Tech Fee", "Probability", "Weighted Value", "Start Month", "End Month"],
    ["Puma Renewal", 500, 0.9, "=G3*H3", new Date("2025-02-01"), new Date("2025-12-31")],
    ["The AA Upsell", 500, 0.5, "=G4*H4", new Date("2025-04-01"), new Date("2025-06-30")],
    ["Deckers AIO Pitch", 500, 0.2, "=G5*H5", new Date("2026-01-01"), new Date("2026-12-31")],
    ["Pipeline Total", "", "", "=SUM(I3:I5)", "", ""]
  ];
  sheet.getRange("F2:K6").setValues(pipelineData);
  sheet.getRange("F2:K2").setFontWeight("bold").setBackground(headerRowColor);
  sheet.getRange("G3:G5").setNumberFormat("$#,##0");
  sheet.getRange("H3:H5").setNumberFormat("0%");
  sheet.getRange("I3:I6").setNumberFormat("$#,##0");
  sheet.getRange("J3:K5").setNumberFormat("mmm-yyyy");
  sheet.getRange("F2:K6").setBorder(true, true, true, true, true, true);
  sheet.getRange("F6:K6").setBackground(revenueColor).setFontWeight("bold");

  // --- SECTION C: FORECAST ---
  sheet.getRange("A9").setValue("3. Cashflow Forecast (Tool Revenue vs Cost)").setTextStyle(titleStyle);
  
  const startRow = 11;
  const labels = [
    "Month", 
    "Tool Costs (Burn Rate)", 
    "Confirmed Tool Revenue", 
    "Pipeline Revenue (Weighted)", 
    "Monthly Surplus / (Deficit)",
    "Cumulative Cashflow (Bank Balance)"
  ];
  sheet.getRange(startRow, 1, labels.length, 1).setValues(labels.map(x => [x])).setFontWeight("bold");
  
  // Generate Timelines
  const colLetters = generateColumnLetters(7, 24); // G to AD (Matches Step 1 Output)
  let formulaData = [[], [], [], [], [], []]; 
  let currentDate = new Date(START_PROJECTION_DATE);
  const toolCosts = getToolCostsByMonth(START_PROJECTION_DATE, END_PROJECTION_DATE);
  const hasToolCosts = Array.isArray(toolCosts) && toolCosts.length === colLetters.length;
  
  colLetters.forEach((col, index) => {
    let year = currentDate.getFullYear();
    let is2026 = year === 2026;
    let burnCell = is2026 ? "$C$5" : "$B$5";
    let myCol = columnToLetter(index + 2); // Start at B
    
    formulaData[0].push(new Date(currentDate)); // Store date, format after set
    const monthCost = hasToolCosts ? toolCosts[index] : `=${burnCell}`;
    formulaData[1].push(monthCost); // Burn (actual transactions if available)
    formulaData[2].push(`=SUM('${DATA_SHEET_NAME}'!${col}2:${col})`); // Revenue (skip header row)
    formulaData[3].push(`=SUMPRODUCT(($J$3:$J$5<=${myCol}${startRow})*($J$3:$J$5<>"")*(($K$3:$K$5="")+($K$3:$K$5>=${myCol}${startRow}))*$I$3:$I$5)`); // Pipeline active window
    formulaData[4].push(`=(${myCol}${startRow+2}+${myCol}${startRow+3})-${myCol}${startRow+1}`); // Surplus
    
    // Cumulative
    if (index === 0) formulaData[5].push(`=${myCol}${startRow+4}`);
    else formulaData[5].push(`=${columnToLetter(index+1)}${startRow+5}+${myCol}${startRow+4}`);
    
    currentDate.setMonth(currentDate.getMonth() + 1);
  });
  
  sheet.getRange(startRow, 2, 6, 24).setValues(formulaData);
  sheet.getRange(startRow, 2, 1, 24).setNumberFormat("MMM-yy");
  sheet.getRange(startRow+1, 2, 5, 24).setNumberFormat("$#,##0");
  
  // Conditional Formatting (Cumulative Row)
  let cumulativeRange = sheet.getRange(startRow+5, 2, 1, 24);
  let rulePositive = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground("#b7e1cd").setRanges([cumulativeRange]).build();
  let ruleNegative = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground("#f4c7c3").setRanges([cumulativeRange]).build();
  sheet.setConditionalFormatRules([rulePositive, ruleNegative]);
  
  sheet.setFrozenColumns(1);
  
  // --- GO LIVE PREDICTOR ---
  // Place go-live indicator outside the pipeline table to avoid circular refs with start/end months
  sheet.getRange("M2").setValue("EARLIEST 2026 UPGRADE MONTH").setFontWeight("bold");
  const goLiveFormula = `=IFERROR(INDEX(N${startRow}:Y${startRow}, MATCH(TRUE, N${startRow+5}:Y${startRow+5} >= 0, 0)), "Insufficient Funds in 2026")`;
  sheet.getRange("M3").setFormula(goLiveFormula)
       .setFontWeight("bold").setFontSize(14).setBackground("#fff2cc")
       .setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");
  sheet.setColumnWidth(1, 220);
}

// ==========================================
// TOOL TRANSACTIONS TAB
// ==========================================
function ensureToolTransactionsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = "Tool Transactions";
  const headers = ["Date", "Tool / Vendor", "Category", "Amount", "Notes", "Month", "Year"];

  let sheet = ss.getSheetByName(SHEET_NAME);
  const isNew = !sheet;
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  // Seed headers without clearing any existing data
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#f1f3f4");
  sheet.setFrozenRows(1);

  // Basic formatting for entry columns
  sheet.getRange("A2:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("D2:D").setNumberFormat("$#,##0.00");
  sheet.getRange("F2:F").setNumberFormat("MMM yyyy");

  // Seed helper formulas on a fresh sheet so month/year auto-populate from Date
    if (isNew && sheet.getLastRow() < 2) {
      sheet.getRange("F2").setFormula('=IF(A2="","",TEXT(A2,"mmm-yyyy"))');
      sheet.getRange("G2").setFormula('=IF(A2="","",YEAR(A2))');
    }

    headers.forEach((_, idx) => sheet.autoResizeColumn(idx + 1));
    SpreadsheetApp.getUi().alert(`"${SHEET_NAME}" tab is ready. Add transactions starting in row 2.`);
}

// Build a month-by-month view of tool costs, similar to the revenue forecast sheet.
function buildToolCostTimeline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SOURCE_SHEET_NAME = "Tool Transactions";
  const TARGET_SHEET_NAME = "Tool Cost Forecast 2025-26";
  const START_DATE = new Date("2025-01-01");
  const END_DATE = new Date("2026-12-31");

  const source = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!source) throw new Error(`Source sheet "${SOURCE_SHEET_NAME}" not found. Use "Create Tool Transactions Tab" first.`);

  const data = source.getDataRange().getValues();
  if (data.length < 2) throw new Error("No tool transaction data found.");

  const monthHeaders = getMonthHeaders(START_DATE, END_DATE);
  const monthIndex = new Map();
  monthHeaders.dates.forEach((d, i) => {
    monthIndex.set(`${d.getFullYear()}-${d.getMonth()}`, i);
  });

  const rowsByVendor = new Map(); // key: vendor string

  for (let i = 1; i < data.length; i++) { // skip header
    const row = data[i];
    const dt = new Date(row[0]);
    const vendor = (row[1] || "").toString().trim() || "Unknown Vendor";
    const category = (row[2] || "").toString().trim();
    const amount = parseCurrency(row[3]);

    if (!amount || amount <= 0) continue;
    if (isNaN(dt) || dt < START_DATE || dt > END_DATE) continue;

    const key = `${dt.getFullYear()}-${dt.getMonth()}`;
    const idx = monthIndex.get(key);
    if (idx == null) continue;

    if (!rowsByVendor.has(vendor)) {
      rowsByVendor.set(vendor, { vendor, category: category || "", total: 0, months: new Array(monthHeaders.labels.length).fill(0) });
    }
    const entry = rowsByVendor.get(vendor);
    entry.total += amount;
    entry.months[idx] += amount;
    if (!entry.category && category) entry.category = category; // keep first non-empty
  }

  const outputRows = Array.from(rowsByVendor.values())
    .sort((a, b) => a.vendor.localeCompare(b.vendor))
    .map(entry => [entry.vendor, entry.category, entry.total, ...entry.months]);

  const headers = ["Vendor", "Category", "Total Cost", ...monthHeaders.labels];

  let target = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!target) target = ss.insertSheet(TARGET_SHEET_NAME);
  else target.clear();

  target.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#EFEFEF");

  if (outputRows.length) {
    target.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    target.getRange(2, 3, outputRows.length, headers.length - 2).setNumberFormat("$#,##0.00");
  }

  target.setFrozenRows(1);
  target.setFrozenColumns(2);
  headers.forEach((_, idx) => target.autoResizeColumn(idx + 1));
  ss.setActiveSheet(target);
}

// ==========================================
// HELPERS
// ==========================================
function getMonthHeaders(startDate, endDate) {
  let dates = [];
  let labels = [];
  let dt = new Date(startDate);
  while (dt <= endDate) {
    dates.push(new Date(dt));
    labels.push(Utilities.formatDate(dt, Session.getScriptTimeZone(), "MMM yyyy"));
    dt.setMonth(dt.getMonth() + 1);
  }
  return { dates: dates, labels: labels };
}

function monthDiff(d1, d2) {
  let months;
  months = (d2.getFullYear() - d1.getFullYear()) * 12;
  months -= d1.getMonth();
  months += d2.getMonth();
  return months <= 0 ? 0 : months + 1;
}

function parseCurrency(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  let clean = value.toString().replace(/[$,£€]/g, '').replace(/,/g, '');
  return parseFloat(clean) || 0;
}

function generateColumnLetters(startIndex, count) {
  let cols = [];
  for (let i = 0; i < count; i++) cols.push(columnToLetter(startIndex + i));
  return cols;
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Sum monthly tool costs from the "Tool Transactions" sheet for the projection window.
// Assumes amounts are entered as positive numbers; they are treated as costs.
function getToolCostsByMonth(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tool Transactions");
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  // Read Date (col A) and Amount (col D)
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const byMonth = {};

  rows.forEach(r => {
    const dt = r[0];
    const amt = Number(r[3]);
    if (!(dt instanceof Date) || isNaN(dt)) return;
    if (!isFinite(amt)) return;

    const key = `${dt.getFullYear()}-${dt.getMonth()}`; // month key
    byMonth[key] = (byMonth[key] || 0) + amt; // positive amounts = costs
  });

  // Expand to full timeline
  let out = [];
  let cursor = new Date(startDate);
  while (cursor <= endDate) {
    const key = `${cursor.getFullYear()}-${cursor.getMonth()}`;
    out.push(byMonth[key] || 0);
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return out;
}
