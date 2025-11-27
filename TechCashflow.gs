/**
 * Tech Cashflow Projection Engine
 * Analyzes client contracts to generate a month-by-month tech revenue forecast.
 */

function generateTechRunRate(options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = opts.suppressUi ? null : SpreadsheetApp.getUi();
  
  // 1. CONFIGURATION: Map these to your actual sheet names
  const SOURCE_SHEET_NAME = "SEO Revenue from Closed Won Opps - Tech Fees Per Client";
  const SOURCE_SPREADSHEET_ID = "1frMM58KKBphwfQewFrNs01Ob4FB3ntX782JVpT8weWU";
  const SOURCE_TAB_NAME = "Tech Fees Per Client";
  const TARGET_SHEET_NAME = "Tech Cashflow Forecast 2025-26";
  
  // Define the timeline we want to project
  const START_PROJECTION_DATE = new Date("2025-01-01");
  const END_PROJECTION_DATE = new Date("2026-12-31");
  
  // 2. FETCH DATA
  // Try to locate or mirror the source tab from the external workbook
  const sourceSheet =
    ss.getSheetByName(SOURCE_SHEET_NAME) ||
    syncExternalSheet_(ss, ui, SOURCE_SHEET_NAME, SOURCE_SPREADSHEET_ID, SOURCE_TAB_NAME, opts);
  
  // Get all data (assuming headers are in row 1)
  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row
  
  // Map column indexes (Zero-based). ADJUST THESE BASED ON YOUR CSV STRUCTURE
  // Based on your file: "SEO Revenue from Closed Won Opps - Tech Fees Per Client.csv"
  const COL_ACCOUNT = 0;       // Account Name
  const COL_START = 2;         // Start Date
  const COL_END = 3;           // End Date
  const COL_TECH_FEE = 15;     // "Tech Fee (Final)" or similar column with the annual/total amount
  const COL_MARKET = 9;        // Market
  
  // 3. PREPARE OUTPUT GRID
  const monthHeaders = getMonthHeaders(START_PROJECTION_DATE, END_PROJECTION_DATE);
  let outputRows = [];
  
  // 4. PROCESS EACH CLIENT
  data.forEach(row => {
    let clientName = row[COL_ACCOUNT];
    let startDate = new Date(row[COL_START]);
    let endDate = new Date(row[COL_END]);
    let totalFee = parseCurrency(row[COL_TECH_FEE]);
    let market = row[COL_MARKET];
    
    // Skip if no fee or invalid dates
    if (!totalFee || totalFee === 0 || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return;
    }
    
    // Calculate Monthly Run Rate (MRR) for this specific contract
    let monthsDuration = monthDiff(startDate, endDate);
    if (monthsDuration < 1) monthsDuration = 1;
    let monthlyRevenue = totalFee / monthsDuration; // Assumes total fee is spread evenly
    
    // Or if the column is ALREADY monthly, use: let monthlyRevenue = totalFee;
    
    // Generate the row for the timeline
    let clientRow = [clientName, market, startDate, endDate, totalFee, monthlyRevenue];
    
    // Loop through projection months and check if client is active
    monthHeaders.dates.forEach(monthDate => {
      if (monthDate >= startDate && monthDate <= endDate) {
        clientRow.push(monthlyRevenue);
      } else {
        clientRow.push(0);
      }
    });
    
    outputRows.push(clientRow);
  });
  
  // 5. RENDER OUTPUT
  let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
  } else {
    targetSheet.clear(); // Clear old data
  }
  
  // Set Headers
  let finalHeaders = ["Client", "Market", "Start Date", "End Date", "Total Fee", "Monthly Revenue", ...monthHeaders.labels];
  targetSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]).setFontWeight("bold").setBackground("#EFEFEF");
  
  // Set Data
  if (outputRows.length > 0) {
    targetSheet.getRange(2, 1, outputRows.length, finalHeaders.length).setValues(outputRows);
  }
  
  // Formatting
  targetSheet.getRange(2, 5, outputRows.length, finalHeaders.length - 4).setNumberFormat("$#,##0.00");
  targetSheet.setFrozenRows(1);
  targetSheet.setFrozenColumns(1);
}

// --- HELPER FUNCTIONS ---

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
  return months <= 0 ? 0 : months + 1; // +1 to include the starting month
}

function parseCurrency(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  // Remove currency symbols and commas
  let clean = value.toString().replace(/[$,£€]/g, '').replace(/,/g, '');
  return parseFloat(clean) || 0;
}

// If the preferred sheet is missing, prompt the user and show available sheet names.
function getSheetOrPrompt_(ss, ui, preferredName, promptMessage) {
  const preferred = ss.getSheetByName(preferredName);
  if (preferred) return preferred;

  const names = ss.getSheets().map(sh => sh.getName()).join(', ');
  const resp = ui.prompt('Sheet not found', `${promptMessage}\n\nPreferred: "${preferredName}"\nAvailable sheets: ${names}`, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) {
    throw new Error('Canceled: sheet selection required.');
  }
  const name = (resp.getResponseText() || '').trim();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" not found. Check the name and try again.`);
  return sh;
}

// Pulls data from another spreadsheet tab into a local sheet (overwrite/replace).
function syncExternalSheet_(ss, ui, localName, externalSpreadsheetId, externalTabName, options) {
  const opts = options || {};
  try {
    const extSs = SpreadsheetApp.openById(externalSpreadsheetId);
    const extSheet = extSs.getSheetByName(externalTabName);
    if (!extSheet) {
      if (ui) ui.alert(`External tab "${externalTabName}" not found in the source workbook.`);
      throw new Error(`External tab "${externalTabName}" missing`);
    }
    const data = extSheet.getDataRange().getValues();
    if (!data.length) throw new Error('External tab is empty');

    let local = ss.getSheetByName(localName);
    if (!local) local = ss.insertSheet(localName);
    local.clearContents();
    local.getRange(1, 1, data.length, data[0].length).setValues(data);
    return local;
  } catch (err) {
    if (ui) ui.alert('Unable to import source tab. Please ensure you have access to the source spreadsheet and the tab name is correct.');
    throw err;
  }
}
