/**
 * Populate the "OnCrawl Monthly URL Budget" column in the Adjustment sheet.
 *
 * The script expects two sheets:
 *   1. "Adjustments" with the columns "Domain", "Override OnCrawl Budget",
 *      "OnCrawl Monthly URL Budget", and "Budget Source".
 *   2. "OnCrawl Stats" with the columns "Domain", "Monthly URL Budget",
 *      "Avg Daily URLs", and (optionally) "Crawl Days In Month".
 *
 * When an override value is present it will be preserved. Otherwise the
 * OnCrawl figures are used to derive a monthly budget. Daily averages are
 * multiplied by the configured number of crawl days (default 30) when a
 * monthly figure is missing.
 */
function populateOncrawlMonthlyBudget(options) {
  var settings = options || {};
  var spreadsheet = settings.spreadsheet || SpreadsheetApp.getActive();
  var adjustmentsSheet = settings.adjustmentsSheetName
    ? spreadsheet.getSheetByName(settings.adjustmentsSheetName)
    : spreadsheet.getSheetByName('Adjustments');
  var statsSheet = settings.statsSheetName
    ? spreadsheet.getSheetByName(settings.statsSheetName)
    : spreadsheet.getSheetByName('OnCrawl Stats');
  var defaultDays = settings.defaultDays || 30;

  if (!adjustmentsSheet) {
    var newName = settings.adjustmentsSheetName || 'Adjustments';
    adjustmentsSheet = spreadsheet.insertSheet(newName);
    // Seed required headers and freeze header row
    var headers = ['Domain','Override OnCrawl Budget','OnCrawl Monthly URL Budget','Budget Source'];
    adjustmentsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    adjustmentsSheet.setFrozenRows(1);
  }
  if (!statsSheet) {
    throw new Error('Could not find the "OnCrawl Stats" sheet.');
  }

  var adjustmentsData = getSheetData(adjustmentsSheet);
  if (adjustmentsData.rows.length === 0) {
    return 0;
  }
  var statsData = getSheetData(statsSheet);

  var statsByDomain = buildStatsByDomain(statsData, defaultDays);

  var domainCol = adjustmentsData.indexes['Domain'];
  var overrideCol = adjustmentsData.indexes['Override OnCrawl Budget'];
  var budgetCol = adjustmentsData.indexes['OnCrawl Monthly URL Budget'];
  var sourceCol = adjustmentsData.indexes['Budget Source'];

  if (domainCol === undefined) {
    throw new Error('The Adjustments sheet must contain a "Domain" column.');
  }
  if (overrideCol === undefined) {
    throw new Error('The Adjustments sheet must contain an "Override OnCrawl Budget" column.');
  }
  if (budgetCol === undefined) {
    throw new Error('The Adjustments sheet must contain an "OnCrawl Monthly URL Budget" column.');
  }
  if (sourceCol === undefined) {
    throw new Error('The Adjustments sheet must contain a "Budget Source" column.');
  }

  var budgets = [];
  var sources = [];
  for (var r = 0; r < adjustmentsData.rows.length; r++) {
    var row = adjustmentsData.rows[r];
    var domain = asString(row[domainCol]);
    var overrideValue = parseNumber(row[overrideCol]);
    var computedBudget = null;
    var source = '';

    if (overrideValue !== null) {
      computedBudget = overrideValue;
      source = 'override';
    } else if (domain && statsByDomain.hasOwnProperty(domain)) {
      var project = statsByDomain[domain];
      computedBudget = project.monthlyBudget;
      source = project.source;
    } else {
      computedBudget = null;
      source = 'missing';
    }

    budgets.push([computedBudget !== null ? computedBudget : '']);
    sources.push([source]);
  }

  var startRow = 2;
  var budgetRange = adjustmentsSheet.getRange(startRow, budgetCol + 1, budgets.length, 1);
  var sourceRange = adjustmentsSheet.getRange(startRow, sourceCol + 1, sources.length, 1);

  budgetRange.setValues(budgets);
  sourceRange.setValues(sources);

  return budgets.length;
}

function getSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    return { headers: [], indexes: {}, rows: [] };
  }
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var headers = values[0];
  var indexes = {};
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (header) {
      indexes[header] = i;
    }
  }
  var rows = values.slice(1);
  return { headers: headers, indexes: indexes, rows: rows };
}

function buildStatsByDomain(statsData, defaultDays) {
  var domainCol = statsData.indexes['Domain'];
  if (domainCol === undefined) {
    throw new Error('The OnCrawl Stats sheet must contain a "Domain" column.');
  }

  var monthlyCol = statsData.indexes['Monthly URL Budget'];
  var dailyCol = statsData.indexes['Avg Daily URLs'];
  var daysCol = statsData.indexes['Crawl Days In Month'];

  var statsByDomain = {};
  for (var i = 0; i < statsData.rows.length; i++) {
    var row = statsData.rows[i];
    var domain = asString(row[domainCol]);
    if (!domain) {
      continue;
    }

    var monthlyBudget = parseNumber(monthlyCol !== undefined ? row[monthlyCol] : null);
    var dailyAverage = parseNumber(dailyCol !== undefined ? row[dailyCol] : null);
    var crawlDays = parseNumber(daysCol !== undefined ? row[daysCol] : null);

    if (monthlyBudget === null && dailyAverage !== null) {
      monthlyBudget = dailyAverage * (crawlDays !== null ? crawlDays : defaultDays);
    }

    statsByDomain[domain] = {
      monthlyBudget: monthlyBudget,
      source: monthlyBudget !== null ? 'oncrawl' : 'missing'
    };
  }
  return statsByDomain;
}

function parseNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  var num = Number(value);
  if (isNaN(num)) {
    return null;
  }
  return num;
}

function asString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}
