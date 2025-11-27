/**
 * Dashboard Builder
 * Generates the "Upgrade Predictor" interface to visualize cashflow vs. burn rate.
 */

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DASHBOARD_NAME = "Upgrade Predictor";
  const DATA_SHEET_NAME = "Tech Cashflow Forecast 2025-26"; // Must match Step 1 output
  
  // 1. CREATE OR RESET SHEET
  let sheet = ss.getSheetByName(DASHBOARD_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(DASHBOARD_NAME);
  }
  
  // 2. SETUP STYLING CONSTANTS
  const headerStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).build();
  const titleStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(14).setForegroundColor("#1a73e8").build();
  const burnColor = "#fce8e6"; // Red tint for costs
  const revenueColor = "#e6f4ea"; // Green tint for revenue
  
  // 3. BUILD SECTION A: THE BURN RATE (Fixed Costs)
  sheet.getRange("A1").setValue("1. Monthly Burn Rate").setTextStyle(titleStyle);
  
  const costData = [
    ["Item", "Monthly Cost", "Notes"],
    ["Current Tech Stack (AccuRanker + Semrush)", 13144, "=(57732+100000)/12"], // Approx fixed cost
    ["Proposed Upgrade (Semrush AIO - 2 Workspaces)", 3000, "Estimated"],
    ["Total Monthly Burn", "=SUM(B3:B4)", ""]
  ];
  
  sheet.getRange("A2:C5").setValues(costData);
  sheet.getRange("A2:C2").setFontWeight("bold").setBackground("#eeeeee");
  sheet.getRange("B3:B5").setNumberFormat("$#,##0");
  sheet.getRange("A2:C5").setBorder(true, true, true, true, true, true);
  sheet.getRange("A5:C5").setBackground(burnColor).setFontWeight("bold"); // Highlight Total Burn
  
  // 4. BUILD SECTION B: THE PIPELINE CALCULATOR ("What-If")
  sheet.getRange("E1").setValue("2. Pipeline Scenario Builder").setTextStyle(titleStyle);
  
  const pipelineHeaders = [["Client Target", "Est. Monthly Fee", "Win Probability", "Weighted Value"]];
  sheet.getRange("E2:H2").setValues(pipelineHeaders).setFontWeight("bold").setBackground("#eeeeee");
  
  // Add some placeholder rows for the user to edit
  const pipelinePlaceholders = [
    ["Puma Renewal (Example)", 500, 0.9, "=F3*G3"],
    ["The AA Upsell (Example)", 500, 0.5, "=F4*G4"],
    ["Deckers AIO Pitch", 500, 0.2, "=F5*G5"],
    ["Total Pipeline Impact", "", "", "=SUM(H3:H5)"]
  ];
  
  sheet.getRange("E3:H6").setValues(pipelinePlaceholders);
  sheet.getRange("F3:F5").setNumberFormat("$#,##0");
  sheet.getRange("G3:G5").setNumberFormat("0%");
  sheet.getRange("H3:H6").setNumberFormat("$#,##0");
  sheet.getRange("E2:H6").setBorder(true, true, true, true, true, true);
  sheet.getRange("E6:H6").setBackground(revenueColor).setFontWeight("bold");
  
  // 5. BUILD SECTION C: THE CASHFLOW TIMELINE
  sheet.getRange("A9").setValue("3. Cashflow Forecast & Go-Live Prediction").setTextStyle(titleStyle);
  
  // We need to map the columns from the Data Sheet ("Tech Cashflow Forecast 2025-26")
  // In Step 1, Jan 2025 is Column G (Index 7), Feb 2025 is H (Index 8), etc.
  
  const startRow = 11;
  const labels = ["Month", "Guaranteed Revenue (from Contracts)", "Pipeline Revenue (Weighted)", "Total Capacity", "Surplus / (Deficit)"];
  
  // Place Row Labels
  sheet.getRange(startRow, 1, labels.length, 1).setValues(labels.map(x => [x])).setFontWeight("bold");
  
  // Generate Formulas for Jan 2025 - Dec 2026 (24 Months)
  // 'Tech Cashflow Forecast 2025-26'!G:G sums the Jan 2025 column
  
  const colLetters = generateColumnLetters(7, 24); // Get G, H, I... for 24 months
  let formulaData = [[], [], [], [], []]; // 5 rows of data
  
  // Row 1: Months (Headers)
  // We'll hardcode dates for simplicity or pull from config. Let's assume standard Jan 25 start.
  let currentDate = new Date("2025-01-01");
  
  colLetters.forEach(col => {
    // 1. Month Header
    formulaData[0].push(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMM-yy"));
    
    // 2. Guaranteed Revenue formula: =SUM('DataSheet'!Col:Col)
    formulaData[1].push(`=SUM('${DATA_SHEET_NAME}'!${col}:${col})`);
    
    // 3. Pipeline Revenue: Fixed reference to our calculator above ($H$6)
    formulaData[2].push(`=$H$6`);
    
    // 4. Total Capacity: Guaranteed + Pipeline
    let myCol = columnToLetter(formulaData[0].length + 1); // B, C, D... relative to this sheet
    formulaData[3].push(`=SUM(${myCol}${startRow+1}:${myCol}${startRow+2})`);
    
    // 5. Surplus: Total Capacity - Total Burn ($B$5)
    formulaData[4].push(`=${myCol}${startRow+3}-$B$5`);
    
    currentDate.setMonth(currentDate.getMonth() + 1);
  });
  
  // Paste the timeline data
  sheet.getRange(startRow, 2, 5, 24).setValues(formulaData);
  
  // Formatting the timeline
  sheet.getRange(startRow+1, 2, 4, 24).setNumberFormat("$#,##0"); // Currency for all numbers
  
  // Conditional Formatting for the "Surplus" row (Row 15)
  // Green if positive, Red if negative
  let surplusRange = sheet.getRange(startRow+4, 2, 1, 24);
  let rulePositive = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#b7e1cd") // Green
    .setRanges([surplusRange])
    .build();
  let ruleNegative = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground("#f4c7c3") // Red
    .setRanges([surplusRange])
    .build();
    
  sheet.setConditionalFormatRules([rulePositive, ruleNegative]);
  
  // 6. THE "GO LIVE" INDICATOR (Big Metric at Top)
  sheet.getRange("J2").setValue("ESTIMATED GO-LIVE DATE").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Formula: Index Match to find the first month where Surplus > 0
  // Note: This is a complex formula to generate via script, simplified logic:
  // "Look at Row 15, find first value > 0, return Row 11 header"
  
  const firstSurplusCell = "B" + (startRow + 4);
  const lastSurplusCell = "Y" + (startRow + 4); // 24 months later
  const firstDateCell = "B" + startRow;
  const lastDateCell = "Y" + startRow;
  
  // XLOOKUP is best if available, otherwise Index/Match
  // =INDEX(Dates, MATCH(TRUE, SurplusRange > 0, 0))
  // Using ArrayFormula for compatibility
  const goLiveFormula = `=IFERROR(INDEX(${firstDateCell}:${lastDateCell}, MATCH(TRUE, ${firstSurplusCell}:${lastSurplusCell} >= 0, 0)), "Insufficient Budget")`;
  
  sheet.getRange("J3").setFormula(goLiveFormula).setFontWeight("bold").setFontSize(16)
       .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setBackground("#fff2cc");
  
  sheet.setColumnWidth(1, 200); // Widen first column for labels
}

// --- HELPER TO GET EXCEL COLUMN LETTERS (G, H, I...) ---
function generateColumnLetters(startIndex, count) {
  let cols = [];
  for (let i = 0; i < count; i++) {
    cols.push(columnToLetter(startIndex + i));
  }
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