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
    ["Client Target", "Est. Tech Fee", "Probability", "Weighted Value"],
    ["Puma Renewal", 500, 0.9, "=G3*H3"],
    ["The AA Upsell", 500, 0.5, "=G4*H4"],
    ["Deckers AIO Pitch", 500, 0.2, "=G5*H5"],
    ["Pipeline Total", "", "", "=SUM(I3:I5)"]
  ];
  sheet.getRange("F2:I6").setValues(pipelineData);
  sheet.getRange("F2:I2").setFontWeight("bold").setBackground(headerRowColor);
  sheet.getRange("G3:G5").setNumberFormat("$#,##0");
  sheet.getRange("H3:H5").setNumberFormat("0%");
  sheet.getRange("I3:I6").setNumberFormat("$#,##0");
  sheet.getRange("F6:I6").setBackground(revenueColor).setFontWeight("bold");

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
  let currentDate = new Date("2025-01-01");
  
  colLetters.forEach((col, index) => {
    let year = currentDate.getFullYear();
    let monthLabel = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMM-yy");
    let is2026 = year === 2026;
    let burnCell = is2026 ? "$C$5" : "$B$5";
    let myCol = columnToLetter(index + 2); // Start at B
    
    formulaData[0].push(monthLabel);
    formulaData[1].push(`=${burnCell}`); // Burn
    formulaData[2].push(`=SUM('${DATA_SHEET_NAME}'!${col}2:${col})`); // Revenue (skip header row)
    formulaData[3].push(`=$I$6`); // Pipeline
    formulaData[4].push(`=(${myCol}${startRow+2}+${myCol}${startRow+3})-${myCol}${startRow+1}`); // Surplus
    
    // Cumulative
    if (index === 0) formulaData[5].push(`=${myCol}${startRow+4}`);
    else formulaData[5].push(`=${columnToLetter(index+1)}${startRow+5}+${myCol}${startRow+4}`);
    
    currentDate.setMonth(currentDate.getMonth() + 1);
  });
  
  sheet.getRange(startRow, 2, 6, 24).setValues(formulaData);
  sheet.getRange(startRow+1, 2, 5, 24).setNumberFormat("$#,##0");
  
  // Conditional Formatting (Cumulative Row)
  let cumulativeRange = sheet.getRange(startRow+5, 2, 1, 24);
  let rulePositive = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground("#b7e1cd").setRanges([cumulativeRange]).build();
  let ruleNegative = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground("#f4c7c3").setRanges([cumulativeRange]).build();
  sheet.setConditionalFormatRules([rulePositive, ruleNegative]);
  
  sheet.setFrozenColumns(1);
  
  // --- GO LIVE PREDICTOR ---
  sheet.getRange("K2").setValue("EARLIEST 2026 UPGRADE MONTH").setFontWeight("bold");
  const goLiveFormula = `=IFERROR(INDEX(N${startRow}:Y${startRow}, MATCH(TRUE, N${startRow+5}:Y${startRow+5} >= 0, 0)), "Insufficient Funds in 2026")`;
  sheet.getRange("K3").setFormula(goLiveFormula)
       .setFontWeight("bold").setFontSize(14).setBackground("#fff2cc")
       .setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");
  sheet.setColumnWidth(1, 220);
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
