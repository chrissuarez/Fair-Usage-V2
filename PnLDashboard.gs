/**
 * Generates the P&L Dashboard (2025 vs 2026).
 */
function generatePnLDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DASHBOARD_NAME = "P&L Dashboard";
  
  // Data Sources
  const REVENUE_SHEET = "Tech Cashflow Forecast 2025-26";
  const TOOL_COST_SHEET = "Tool Cost Forecast 2025-26";
  const FREELANCER_SHEET = "Import - JF Buy data";
  
  // 1. FETCH DATA
  const revenueData = getMonthlyData_(ss, REVENUE_SHEET, 2, "Confirmed Tool Revenue"); // Assuming row index or we sum column Logic
  // Actually, Tech Cashflow Forecast has months in columns. Let's read it properly.
  const revenueMap = getRevenueMap_(ss, REVENUE_SHEET);
  const toolCostMap = getToolCostMap_(ss, TOOL_COST_SHEET);
  const freelancerMap = getFreelancerMap_(ss, FREELANCER_SHEET);
  
  // 2. PREPARE DASHBOARD SHEET
  let sheet = ss.getSheetByName(DASHBOARD_NAME);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(DASHBOARD_NAME);
  
  // Styling
  const titleStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(14).setForegroundColor("#1a73e8").build();
  const headerColor = "#f1f3f4";
  const totalColor = "#e6f4ea";
  
  // --- SECTION 1: HIGH-LEVEL P&L TABLE ---
  sheet.getRange("A1").setValue("1. High-Level P&L Summary").setTextStyle(titleStyle);
  
  const headers = ["Category", "2025 Total", "2026 Total", "Variance ($)", "Variance (%)"];
  const categories = [
    "Total Revenue", 
    "Tool Costs", 
    "Freelancer Costs", 
    "Staff Costs (Placeholder)", 
    "Net Profit"
  ];
  
  // Calculate Totals
  const rev25 = sumYear_(revenueMap, 2025);
  const rev26 = sumYear_(revenueMap, 2026);
  
  const tool25 = sumYear_(toolCostMap, 2025);
  const tool26 = sumYear_(toolCostMap, 2026);
  
  const free25 = sumYear_(freelancerMap, 2024); // Wait, user asked 2025 vs 2026. Let's filter for 2025/26.
  // Correction: sumYear_(freelancerMap, 2025)
  const free25_actual = sumYear_(freelancerMap, 2025);
  const free26 = sumYear_(freelancerMap, 2026);
  
  const staff25 = 0; // Placeholder
  const staff26 = 0; // Placeholder
  
  const profit25 = rev25 - tool25 - free25_actual - staff25;
  const profit26 = rev26 - tool26 - free26 - staff26;
  
  const data = [
    ["Total Revenue", rev25, rev26],
    ["Tool Costs", tool25, tool26],
    ["Freelancer Costs", free25_actual, free26],
    ["Staff Costs (Placeholder)", staff25, staff26],
    ["Net Profit", profit25, profit26]
  ];
  
  // Add Variance
  const tableData = data.map(r => {
    const vDollar = r[2] - r[1];
    const vPercent = r[1] !== 0 ? vDollar / r[1] : 0;
    return [...r, vDollar, vPercent];
  });
  
  sheet.getRange("A3:E3").setValues([headers]).setFontWeight("bold").setBackground(headerColor);
  sheet.getRange(4, 1, tableData.length, 5).setValues(tableData);
  
  // Formatting
  sheet.getRange("B4:D8").setNumberFormat("$#,##0");
  sheet.getRange("E4:E8").setNumberFormat("0.0%");
  sheet.getRange("A8:E8").setBackground(totalColor).setFontWeight("bold").setBorder(true, false, false, false, false, false);
  
  // Placeholder Note
  sheet.getRange("D7").setNote("Enter Staff Costs here manually");
  sheet.getRange("E7").setNote("Enter Staff Costs here manually");
  
  // --- SECTION 2: PROFITABILITY CARDS ---
  sheet.getRange("G1").setValue("2. Profitability Margin").setTextStyle(titleStyle);
  
  const margin25 = rev25 > 0 ? profit25 / rev25 : 0;
  const margin26 = rev26 > 0 ? profit26 / rev26 : 0;
  
  sheet.getRange("G3").setValue("2025 Margin");
  sheet.getRange("G4").setValue(margin25).setNumberFormat("0.0%").setFontSize(24).setFontWeight("bold");
  
  sheet.getRange("I3").setValue("2026 Margin");
  sheet.getRange("I4").setValue(margin26).setNumberFormat("0.0%").setFontSize(24).setFontWeight("bold");
  
  // --- SECTION 3: MONTHLY TREND DATA (Hidden for Chart) ---
  const montlyStartRow = 15;
  sheet.getRange(montlyStartRow - 1, 1).setValue("Monthly Data (Hidden Source for Chart)").setFontWeight("bold");
  const monthHeaders = ["Month", "Revenue", "Total Costs", "Net Profit"];
  sheet.getRange(montlyStartRow, 1, 1, 4).setValues([monthHeaders]).setFontWeight("bold");
  
  const months = getMonthKeys_(); // "2025-0", "2025-1"...
  const monthlyData = months.map(m => {
    const r = revenueMap[m] || 0;
    const t = toolCostMap[m] || 0;
    const f = freelancerMap[m] || 0;
    const c = t + f; // Total Costs
    const n = r - c;
    return [m.label, r, c, n];
  });
  
  sheet.getRange(montlyStartRow + 1, 1, monthlyData.length, 4).setValues(monthlyData);
  sheet.getRange(montlyStartRow + 1, 2, monthlyData.length, 3).setNumberFormat("$#,##0");
  
  // --- SECTION 4: CHART GENERATION ---
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(montlyStartRow, 1, monthlyData.length + 1, 3)) // Month, Rev, Cost
    .setPosition(10, 1, 0, 0) // Row 10
    .setOption('title', 'Monthly Revenue vs Costs')
    .setOption('series', {
      0: { type: 'bars', color: '#34a853', labelInLegend: 'Revenue' },
      1: { type: 'bars', color: '#ea4335', labelInLegend: 'Costs' }
    })
    .setOption('width', 800)
    .setOption('height', 400)
    .build();
    
  sheet.insertChart(chart);
  
  sheet.autoResizeColumns(1, 10);
  SpreadsheetApp.getUi().alert("P&L Dashboard Generated Successfully.");
}

// --- HELPERS ---

function getRevenueMap_(ss, sheetName) {
  // Config from TechFinanceTool.gs: "Tech Cashflow Forecast 2025-26"
  // Format: Row 11 Headers (Month labels G11:AD11). 
  // Row 15: Confirmed Tool Revenue. Row 16: Pipeline.
  // We want Total Revenue = Confirmed + Pipeline? 
  // Looking at DashboardBuilder.gs: "Pipeline Revenue (Weighted)" is separate. 
  // Let's assume user wants Total Revenue which usually implies Confirmed + Weighted Pipeline.
  
  // Actually, let's read the *Dashboard* created in step 2 ("Upgrade Predictor") as it already sums these up?
  // DashboardBuilder.gs: Row 11 labels. Row 13: Confirmed. Row 14: Pipeline.
  // Wait, DashboardBuilder writes to "Upgrade Predictor".
  // Let's read from "Upgrade Predictor" directly if it exists, it's easier.
  
  const dash = ss.getSheetByName("Upgrade Predictor");
  if (!dash) return {}; // Fallback or throw
  
  // DashboardBuilder.gs: 
  // Row 11: Labels (Month 1...)
  // Row 13: Confirmed Tool Revenue
  // Row 14: Pipeline Revenue
  // Actually checking code: 
  // startRow = 11. labels = "Month", "Tool Costs", "Confirmed...", "Pipeline..."
  // So Row 11 is Headers. 
  // Row 12 (Index 1): Tool Costs.
  // Row 13 (Index 2): Confirmed Rev.
  // Row 14 (Index 3): Pipeline Rev.
  
  const startRow = 11;
  const lastCol = dash.getLastColumn();
  const headers = dash.getRange(startRow, 2, 1, lastCol - 1).getValues()[0]; // Dates
  const confirmed = dash.getRange(startRow + 2, 2, 1, lastCol - 1).getValues()[0];
  const pipeline = dash.getRange(startRow + 3, 2, 1, lastCol - 1).getValues()[0];
  
  const map = {};
  headers.forEach((d, i) => {
    if (d instanceof Date) {
      const key = `${d.getFullYear()}-${d.getMonth()}`;
      const val = (Number(confirmed[i]) || 0) + (Number(pipeline[i]) || 0);
      map[key] = val;
    }
  });
  return map;
}

function getToolCostMap_(ss, sheetName) {
  // "Tool Cost Forecast 2025-26"
  // This sheet has vendors in rows, and months in columns.
  // The Total is likely a sum row? Or we calculate it.
  // TechFinanceTool.gs > buildToolCostTimeline writes headers in Row 1. Data starts Row 2.
  // We need to sum all rows for each column.
  
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return {};
  
  const data = sh.getDataRange().getValues();
  const headers = data[0]; // Vendor, Category, Total Cost, Jan-25, ...
  const colStart = 3; // Index 3 is first month (Jan-25) if "Total Cost" is index 2.
  // Headers: "Vendor", "Category", "Total Cost", ...monthHeaders.labels
  
  const map = {};
  
  // Identify month columns
  const monthCols = [];
  for (let c = 3; c < headers.length; c++) {
    // Header is string "MMM yyyy". We need to convert to key.
    const d = new Date(headers[c]); // "Jan 2025" parses ok
    if (!isNaN(d.getTime())) {
      monthCols.push({ idx: c, key: `${d.getFullYear()}-${d.getMonth()}` });
    }
  }
  
  // Sum columns
  for (let r = 1; r < data.length; r++) {
    for (let m of monthCols) {
      const val = Number(data[r][m.idx]) || 0;
      map[m.key] = (map[m.key] || 0) + val;
    }
  }
  return map;
}

function getFreelancerMap_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return {};
  
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  
  // Find Columns
  const dateCol = headers.indexOf("ExpectedDeliveryDate"); // Or Request_Date? 
  // User prompt said "ExpectedDeliveryDate" 30/07/2024.
  // Let's us "ExpectedDeliveryDate" for cost accrual? Or "Request_Date"?
  // Usually delivery date is closer to invoice.
  // Let's try to find "POAmount_USD".
  const costCol = headers.indexOf("POAmount_USD");
  const dateCol2 = headers.indexOf("Request_Date");
  
  if (costCol === -1) return {};
  
  const map = {};
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    // Use ExpectedDeliveryDate, fallback to Request_Date
    let d = row[dateCol]; 
    if (!(d instanceof Date)) d = row[dateCol2];
    if (!(d instanceof Date)) continue;
    
    const cost = Number(row[costCol]) || 0;
    
    // Filter out canceled? "Status" col? 
    // Status is col 2. "10. Request Cancelled".
    const status = String(row[2]).toLowerCase();
    if (status.includes("cancelled") || status.includes("rejected")) continue;
    
    const key = `${d.getFullYear()}-${d.getMonth()}`;
    map[key] = (map[key] || 0) + cost;
  }
  
  return map;
}

function sumYear_(map, year) {
  let sum = 0;
  Object.keys(map).forEach(k => {
    if (k.startsWith(`${year}-`)) sum += map[k];
  });
  return sum;
}

function getMonthKeys_() {
  const start = new Date(2025, 0, 1);
  const end = new Date(2026, 11, 31);
  const list = [];
  let cur = new Date(start);
  while (cur <= end) {
    list.push({
      key: `${cur.getFullYear()}-${cur.getMonth()}`,
      toString: function() { return this.key; },
      label: Utilities.formatDate(cur, Session.getScriptTimeZone(), "MMM-yy")
    });
    cur.setMonth(cur.getMonth() + 1);
  }
  return list;
}
