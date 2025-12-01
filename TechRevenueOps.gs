/************************************************************
 * Tech Revenue Ops pipeline
 * Implements the requirements from Tech_Revenue_Ops_System_Requirements.md:
 * - Shadow table ingestion with overrides
 * - Revenue categorization and monthly spread
 * - Policy health check with target tiers/fees
 * - Portfolio Health + Renewal Radar dashboards
 ************************************************************/

const REVENUE_SOURCE = 'Estimate';
const TECHFEE_SOURCE = 'TechFee';

const SOURCE_CONFIG = {
  Estimate: {
    rawSheet: `${REVENUE_SOURCE}_RAW_Data_Import`,
    overrideSheet: `${REVENUE_SOURCE}_MANUAL_Overrides`,
    masterSheet: `${REVENUE_SOURCE}_MASTER_Ledger`,
    requiredHeaders: [
      'Estimate Name',
      'Resource Role',
      'Pricing Region: Region Name',
      'Hours',
      'Requested Bill Rate (converted) Currency',
      'Requested Bill Rate (converted)',
      'Start Date',
      'End Date',
      'Estimate ID',
      'Account: Account Name',
      'Opportunity: Opportunity Name'
    ],
    fieldMap: {
      estimateName: 'Estimate Name',
      role: 'Resource Role',
      region: 'Pricing Region: Region Name',
      hours: 'Hours',
      currency: 'Requested Bill Rate (converted) Currency',
      billRate: 'Requested Bill Rate (converted)',
      startDate: 'Start Date',
      endDate: 'End Date',
      estimateId: 'Estimate ID',
      account: 'Account: Account Name',
      opportunity: 'Opportunity: Opportunity Name'
    },
    amountField: 'Revenue'
  },
  TechFee: {
    rawSheet: `${TECHFEE_SOURCE}_RAW_Data_Import`,
    overrideSheet: `${TECHFEE_SOURCE}_MANUAL_Overrides`,
    masterSheet: `${TECHFEE_SOURCE}_MASTER_Ledger`,
    requiredHeaders: [
      'Close Date',
      'Account Name',
      'Opportunity Name',
      'Opportunity ID',
      'Product Name',
      'Amount (Net Revenue) (converted) Currency',
      'Amount (Net Revenue) (converted)',
      'Sales Price (Net Revenue) (converted) Currency',
      'Sales Price (Net Revenue) (converted)',
      'Stage'
    ],
    fieldMap: {
      closeDate: 'Close Date',
      account: 'Account Name',
      opportunity: 'Opportunity Name',
      opportunityUid: 'Opportunity ID',
      productName: 'Product Name',
      productAmount: 'Sales Price (Net Revenue) (converted)', // User requested Column I
      currency: 'Amount (Net Revenue) (converted) Currency',
      salesPrice: 'Sales Price (Net Revenue) (converted)',
      stage: 'Stage'
    },
    amountField: 'Product Amount'
  },
  Projects: {
    rawSheet: 'Projects_RAW_Data_Import',
    externalId: '1U6CobpJhF4705cRhRYdoYO67vlIIuAG0tDLFF_242CE',
    tabName: 'Import',
    requiredHeaders: ['Project: Project Name', 'Account', 'Start Date', 'End Date', 'Stage', 'Opportunity'],
    fieldMap: {
      project: 'Project: Project Name',
      account: 'Account',
      startDate: 'Start Date',
      endDate: 'End Date',
      stage: 'Stage',
      opportunity: 'Opportunity'
    }
  }
};

/************************************************************
 * Web endpoints for Revenue Ops UI
 ************************************************************/
function getRevenueOpsDashboardData_Web() {
  return getRevenueOpsDashboardData_();
}

function updateOpportunityRevenue_Web(payload) {
  const uid = payload && payload.uid;
  const amount = payload && Number(payload.amount);
  const reason = (payload && payload.reason) || 'Web edit';
  if (!uid || !isFinite(amount)) throw new Error('Missing uid or amount');
  appendOverride_('Estimate_MANUAL_Overrides', uid, 'Revenue', amount, reason);
  refreshRevenueOpsPipeline();
  return 'Revenue override saved and pipeline refreshed.';
}

function addTechFeeForOpportunity_Web(payload) {
  const uid = payload && payload.uid;
  const amount = payload && Number(payload.amount);
  const reason = (payload && payload.reason) || 'Web tech fee add';
  if (!uid || !isFinite(amount)) throw new Error('Missing uid or amount');

  appendOverride_('TechFee_MANUAL_Overrides', uid, 'Product Amount', amount, reason);

  const master = readMasterRowByUid_(uid, false);
  if (master) {
    const newRevenue = Math.max(0, (master.totalUsd || 0) - amount);
    appendOverride_('Estimate_MANUAL_Overrides', uid, 'Revenue', newRevenue, 'Auto adjust after tech fee add');
  }

  refreshRevenueOpsPipeline();
  return 'Tech fee added, revenue adjusted, and pipeline refreshed.';
}

function getRevenueOpsDashboardData_() {
  const ss = SpreadsheetApp.getActive();
  const masterSh = ss.getSheetByName('MASTER_Ledger');
  const portfolioSh = ss.getSheetByName('Portfolio Health');
  const renewalSh = ss.getSheetByName('Renewal Radar');
  if (!masterSh) throw new Error('MASTER_Ledger not found. Run Refresh Revenue Ops Pipeline.');
  if (masterSh.getLastRow() < 2) throw new Error('MASTER_Ledger is empty. Paste data into RAW tabs and run Refresh Revenue Ops Pipeline.');

  const masterValues = masterSh.getDataRange().getValues();
  const headers = masterValues.shift() || [];
  const idx = {};
  headers.forEach((h, i) => idx[safeStr_(h).toLowerCase()] = i);
  const required = ['account', 'opportunity name', 'opportunity_uid', 'capability', 'start date', 'end date', 'total_usd'];
  const missing = required.filter(k => idx[k] === undefined);
  if (missing.length) {
    throw new Error('MASTER_Ledger missing columns: ' + missing.join(', '));
  }

  const accountMap = {};
  const opportunities = [];
  const today = new Date();
  masterValues.forEach(r => {
    const account = safeStr_(r[idx['account']]);
    const opp = safeStr_(r[idx['opportunity name']]);
    const uid = safeStr_(r[idx['opportunity_uid']]);
    const capability = safeStr_(r[idx['capability']]);
    const start = r[idx['start date']] instanceof Date ? r[idx['start date']] : parseDate_(r[idx['start date']]);
    const end = r[idx['end date']] instanceof Date ? r[idx['end date']] : parseDate_(r[idx['end date']]);
    const totalUsd = toNumber_(r[idx['total_usd']]);
    const monthly = {};
    headers.forEach((h, col) => {
      let key = '';
      if (h instanceof Date) {
        key = monthKey_(h);
      } else {
        const hs = safeStr_(h);
        key = /^\d{4}-\d{2}$/.test(hs) ? hs : '';
      }
      if (key) monthly[key] = toNumber_(r[col]);
    });

    const capLower = capability.toLowerCase();
    const isTechFee = capLower.indexOf('tech fee') !== -1
      || capLower.indexOf('tech & tools') !== -1
      || capLower.indexOf('tech') !== -1
      || capLower.indexOf('tool') !== -1;
    const monthlySum = sumValues_(Object.values(monthly));
    const monthsCount = Object.keys(monthly).length || 1;
    const annual = (monthlySum / monthsCount) * 12;

    accountMap[account] = accountMap[account] || {
      account,
      seoAnnual: 0,
      techAnnual: 0,
      opportunities: [],
      earliestEnd: end
    };
    if (isTechFee) accountMap[account].techAnnual += annual;
    else accountMap[account].seoAnnual += annual;
    if (end && (!accountMap[account].earliestEnd || end < accountMap[account].earliestEnd)) {
      accountMap[account].earliestEnd = end;
    }

    opportunities.push({
      uid,
      account,
      opportunity: opp,
      capability,
      startDate: start,
      endDate: end,
      totalUsd,
      techFee: isTechFee,
      annual
    });
    accountMap[account].opportunities.push(uid);
  });

  const health = {};
  if (portfolioSh) {
    const vals = portfolioSh.getDataRange().getValues();
    vals.shift();
    vals.forEach(r => {
      const acc = safeStr_(r[0]);
      health[acc] = {
        tier: safeStr_(r[2]),
        targetFee: toNumber_(r[3]),
        actualFee: toNumber_(r[4]),
        variance: toNumber_(r[5]),
        status: safeStr_(r[6]),
        plan: safeStr_(r[7])
      };
    });
  }

  const accounts = Object.keys(accountMap).map(acc => {
    const h = health[acc] || {};
    return Object.assign({}, accountMap[acc], {
      tier: h.tier || '',
      targetFee: h.targetFee || 0,
      actualFee: h.actualFee || accountMap[acc].techAnnual || 0,
      variance: h.variance || ((h.targetFee || 0) - (h.actualFee || accountMap[acc].techAnnual || 0)),
      status: h.status || '',
      plan: h.plan || ''
    });
  }).filter(a => a.seoAnnual > 0 || a.techAnnual > 0);

  let renewals = [];
  if (renewalSh && renewalSh.getLastRow() > 1) {
    const vals = renewalSh.getDataRange().getValues();
    vals.shift();
    renewals = vals.map(r => ({
      account: safeStr_(r[0]),
      revenue: toNumber_(r[1]),
      tier: safeStr_(r[2]),
      targetFee: toNumber_(r[3]),
      actualFee: toNumber_(r[4]),
      status: safeStr_(r[5]),
      plan: safeStr_(r[6]),
      renewal: r[7]
    }));
  } else {
    const cutoff = new Date(today.getTime() + 90 * 86400000);
    renewals = accounts.filter(a => a.earliestEnd && a.earliestEnd <= cutoff).map(a => ({
      account: a.account,
      revenue: a.seoAnnual,
      tier: a.tier,
      targetFee: a.targetFee,
      actualFee: a.actualFee,
      status: a.status,
      plan: a.plan,
      renewal: a.earliestEnd
    }));
  }

  return { accounts, opportunities, renewals, debug: { rows: masterValues.length, headers } };
}

/************************************************************
 * Helpers for overrides and master lookups
 ************************************************************/
function appendOverride_(sheetName, uid, field, newValue, reason) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Missing override sheet ${sheetName}`);
  const row = sh.getLastRow() + 1;
  const vals = [[uid, field, newValue, reason || '', Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : 'web', new Date()]];
  sh.getRange(row, 1, 1, 6).setValues(vals);
}

function readMasterRowByUid_(uid, techFeeOnly) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('MASTER_Ledger');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idx = {};
  headers.forEach((h, i) => idx[safeStr_(h).toLowerCase()] = i);
  let fallback = null;
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (safeStr_(r[idx['opportunity_uid']]) === uid) {
      const cap = safeStr_(r[idx['capability']]);
      const isTech = cap.toLowerCase().indexOf('tech fee') !== -1;
      const obj = {
        totalUsd: toNumber_(r[idx['total_usd']]),
        capability: cap
      };
      if (techFeeOnly) {
        if (isTech) return obj;
      } else {
        if (!isTech) return obj;
        fallback = fallback || obj;
      }
    }
  }
  return fallback;
}

/**
 * Main entry to rebuild master ledger + dashboards end-to-end.
 */
function refreshRevenueOpsPipeline() {
  const cfg = ensureRevenueOpsConfigTabs();
  ensureShadowTables_();
  importProjectsData(); // Import fresh project data

  const estimate = rebuildMasterForSource_(SOURCE_CONFIG.Estimate, cfg);

  // Build date lookup map from Estimate rows (Opportunity Name -> {start, end})
  const dateLookup = {};
  estimate.rows.forEach(r => {
    if (r.Opportunity_Name && r.Start_Date) {
      dateLookup[r.Opportunity_Name] = { start: r.Start_Date, end: r.End_Date };
    }
  });
  
  // Build Project lookup map (Opportunity -> {start, end, name})
  const ss = SpreadsheetApp.getActive();
  const projSh = ss.getSheetByName(SOURCE_CONFIG.Projects.rawSheet);
  const projectLookup = {};
  if (projSh && projSh.getLastRow() > 1) {
    const data = projSh.getDataRange().getValues();
    const headers = data.shift();
    const hMap = headers.reduce((acc, h, i) => { acc[h] = i; return acc; }, {});
    const fm = SOURCE_CONFIG.Projects.fieldMap;
    
    data.forEach(r => {
      const acc = safeStr_(r[hMap[fm.account]]);
      const start = parseDate_(r[hMap[fm.startDate]]);
      const end = parseDate_(r[hMap[fm.endDate]]);
      const name = safeStr_(r[hMap[fm.project]]);
      const opp = safeStr_(r[hMap[fm.opportunity]]);
      
      // Filter: Ignore "OPP_" projects
      if (name.toUpperCase().startsWith('OPP_')) return;

      // Key by Opportunity Name
      if (opp && start && end) {
        projectLookup[opp] = { start, end, name };
      }
    });
  }

  const techFee = rebuildMasterForSource_(SOURCE_CONFIG.TechFee, cfg, dateLookup, projectLookup);

  buildCombinedLedger_(estimate, techFee);
  buildPortfolioHealth_(estimate, techFee, cfg);
  buildRenewalRadar_(estimate, techFee, cfg);

  return 'Revenue Ops pipeline refreshed (master ledger + dashboards).';
}

/**
 * Create/read config tabs: Config_Currency, Config_Tech_SKU_Pricing, Config_Params, Import_Log.
 */
function ensureRevenueOpsConfigTabs() {
  const ss = SpreadsheetApp.getActive();

  const currencySh = getOrCreateSheet_(ss, 'Config_Currency');
  if (currencySh.getLastRow() < 2) {
    currencySh.clear();
    currencySh.getRange(1, 1, 1, 3).setValues([['Currency', 'Rate_to_USD', 'Last_Updated']]).setFontWeight('bold');
    currencySh.getRange(2, 1, 3, 3).setValues([
      ['USD', 1, new Date()],
      ['GBP', 1.26, new Date()],
      ['EUR', 1.09, new Date()]
    ]);
    currencySh.setFrozenRows(1);
  }

  const skuSh = getOrCreateSheet_(ss, 'Config_Tech_SKU_Pricing');
  if (skuSh.getLastRow() < 2) {
    skuSh.clear();
    skuSh.getRange(1, 1, 1, 5).setValues([['Tier', 'SKU_Name', 'Monthly_USD', 'Annual_USD', 'Notes']]).setFontWeight('bold');
    skuSh.getRange(2, 1, 4, 5).setValues([
      ['Tier A', 'Tech Pro', 1000, 12000, 'Enterprise'],
      ['Tier B', 'Tech Pro', 1000, 12000, 'Upper mid'],
      ['Tier C', 'Tech Starter', 500, 6000, 'Mid'],
      ['Tier D', 'Tech Starter', 200, 2400, 'Entry']
    ]);
    skuSh.setFrozenRows(1);
  }

  const paramSh = getOrCreateSheet_(ss, 'Config_Params');
  if (paramSh.getLastRow() < 2) {
    paramSh.clear();
    paramSh.getRange(1, 1, 1, 3).setValues([['Param', 'Value', 'Notes']]).setFontWeight('bold');
    paramSh.getRange(2, 1, 3, 3).setValues([
      ['Today', new Date(), 'Anchor date for renewals'],
      ['Renewal_Lookahead_Days', 90, 'Renewals within N days show in Renewal Radar'],
      ['Partial_Month_Mode', 'PRORATE', 'PRORATE or SIMPLE']
    ]);
    paramSh.setFrozenRows(1);
  }

  const logSh = getOrCreateSheet_(ss, 'Import_Log');
  if (logSh.getLastRow() < 1) {
    logSh.clear();
    logSh.getRange(1, 1, 1, 7).setValues([['Source', 'Imported_By', 'Timestamp', 'Row_Count', 'File_Name', 'Validation_Status', 'Errors']]).setFontWeight('bold');
    logSh.setFrozenRows(1);
  }

  return {
    currencyMap: readCurrencyMap_(currencySh),
    pricing: readSkuPricing_(skuSh),
    params: readParams_(paramSh)
  };
}

/**
 * Create raw/override/master sheets for both sources if missing.
 */
function ensureShadowTables_() {
  [SOURCE_CONFIG.Estimate, SOURCE_CONFIG.TechFee].forEach(cfg => {
    const ss = SpreadsheetApp.getActive();
    const raw = getOrCreateSheet_(ss, cfg.rawSheet);
    raw.getRange(1, 1, 1, cfg.requiredHeaders.length).setValues([cfg.requiredHeaders]).setFontWeight('bold');

    const ov = getOrCreateSheet_(ss, cfg.overrideSheet);
    if (ov.getLastRow() < 1) {
      ov.getRange(1, 1, 1, 6).setValues([['Opportunity_UID', 'Field to Override', 'New Value', 'Reason', 'Entered By', 'Timestamp']]).setFontWeight('bold');
    }
    const master = getOrCreateSheet_(ss, cfg.masterSheet);
    master.getRange(1, 1, 1, cfg.requiredHeaders.length).setValues([cfg.requiredHeaders]).setFontWeight('bold');
  });
  ensureProjectShadowTables_();
}

function ensureProjectShadowTables_() {
  const ss = SpreadsheetApp.getActive();
  const cfg = SOURCE_CONFIG.Projects;
  const raw = getOrCreateSheet_(ss, cfg.rawSheet);
  if (raw.getLastRow() < 1) {
    raw.getRange(1, 1, 1, cfg.requiredHeaders.length).setValues([cfg.requiredHeaders]).setFontWeight('bold');
  }
}

function importProjectsData() {
  const ss = SpreadsheetApp.getActive();
  const cfg = SOURCE_CONFIG.Projects;
  const rawSh = getOrCreateSheet_(ss, cfg.rawSheet);
  
  try {
    const sourceSs = SpreadsheetApp.openById(cfg.externalId);
    const sourceSh = sourceSs.getSheetByName(cfg.tabName);
    if (!sourceSh) throw new Error(`Tab "${cfg.tabName}" not found in external sheet.`);
    
    const data = sourceSh.getDataRange().getValues();
    if (data.length < 2) return 'No data found in external Projects sheet.';
    
    // Clear existing data but keep headers
    if (rawSh.getLastRow() > 1) {
      rawSh.getRange(2, 1, rawSh.getLastRow() - 1, rawSh.getLastColumn()).clearContent();
    }
    
    // Write new data (skip header row from source if we assume row 1 is header)
    // We'll write all data including headers to be safe, or just data?
    // Let's write data starting from row 2, assuming source has headers
    const rows = data.slice(1);
    if (rows.length > 0) {
      rawSh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    return `Imported ${rows.length} projects.`;
  } catch (e) {
    console.error('Error importing projects:', e);
    throw new Error(`Failed to import projects: ${e.message}`);
  }
}

/**
 * UI entry: ensure raw/override/master tabs exist for both sources.
 */
function EnsureRevenueOpsShadowTables() {
  ensureShadowTables_();
  return 'Revenue Ops shadow tables ensured.';
}

/**
 * Rebuild master sheet for a single source, applying overrides, currency conversion, categorization, and monthly spreads.
 */
function rebuildMasterForSource_(cfg, globalCfg, dateLookup, projectLookup) {
  const ss = SpreadsheetApp.getActive();
  const rawSh = ss.getSheetByName(cfg.rawSheet);
  const ovSh = ss.getSheetByName(cfg.overrideSheet);
  const masterSh = ss.getSheetByName(cfg.masterSheet);
  if (!rawSh || !ovSh || !masterSh) throw new Error(`Missing shadow table for ${cfg.rawSheet}`);

  const headers = rawSh.getRange(1, 1, 1, rawSh.getLastColumn()).getDisplayValues()[0];
  const rawRows = rawSh.getLastRow() > 1 ? rawSh.getRange(2, 1, rawSh.getLastRow() - 1, headers.length).getValues() : [];
  const overrides = readOverrides_(ovSh);

  const monthsSet = new Set();
  const masterRows = [];
  const headerMap = headers.reduce((acc, h, idx) => { acc[h] = idx; return acc; }, {});

  rawRows.forEach(raw => {
    const get = name => raw[headerMap[name]];
    let uid = '';
    let account = '';
    let oppName = '';
    let startDate = null;
    let endDate = null;
    let currency = 'USD';
    let totalAmount = 0;
    let capability = '';
    let techFeePaying = false;
    let debugInfo = [];

    if (cfg === SOURCE_CONFIG.Estimate) {
      const fm = cfg.fieldMap;
      account = safeStr_(get(fm.account));
      oppName = safeStr_(get(fm.opportunity));
      const estId = safeStr_(get(fm.estimateId));
      uid = estId || generateOpportunityUid_(account, oppName);
      startDate = parseDate_(get(fm.startDate));
      endDate = parseDate_(get(fm.endDate));
      currency = safeStr_(get(fm.currency)) || 'USD';
      const hours = toNumber_(get(fm.hours));
      const billRate = toNumber_(get(fm.billRate));
      totalAmount = hours * billRate;
      capability = categorizeRevenue(get(fm.role));
      techFeePaying = false;
    } else {
      const fm = cfg.fieldMap;
      uid = safeStr_(get(fm.opportunityUid)) || generateOpportunityUid_(safeStr_(get(fm.account)), safeStr_(get(fm.opportunity)));
      account = safeStr_(get(fm.account));
      oppName = safeStr_(get(fm.opportunity));
      
      // 1. Try Project Lookup (by Opportunity Name)
      if (projectLookup && projectLookup[oppName]) {
        startDate = projectLookup[oppName].start;
        endDate = projectLookup[oppName].end;
        debugInfo.push(`Dates from Project: ${projectLookup[oppName].name}`);
      }
      
      // 2. Try Opportunity Lookup (by Name) - Fallback to Estimate dates if no Project match
      if (!startDate && dateLookup && dateLookup[oppName]) {
        startDate = dateLookup[oppName].start;
        endDate = dateLookup[oppName].end;
        debugInfo.push('Dates from Estimate Opp');
      }

      // 3. Fallback to Close Date + 1 Year
      if (!startDate) {
        if (fm.startDate) startDate = parseDate_(get(fm.startDate));
        if (!startDate && fm.closeDate) startDate = parseDate_(get(fm.closeDate));
        
        if (startDate) {
           if (fm.endDate) endDate = parseDate_(get(fm.endDate));
           if (!endDate) {
             const d = new Date(startDate);
             d.setFullYear(d.getFullYear() + 1);
             d.setDate(d.getDate() - 1);
             endDate = d;
             debugInfo.push('Dates Defaulted (1yr)');
           }
        }
      }
      
      currency = safeStr_(get(fm.currency)) || 'USD';
      totalAmount = toNumber_(get(fm.productAmount));
      capability = safeStr_(get(fm.productName)) || 'Tech Fee';
      
      const stage = safeStr_(get(fm.stage)).trim().toLowerCase();
      // User request: All Tech Fee rows are paying
      techFeePaying = true; 
      if (!techFeePaying) debugInfo.push(`Not Paying: Stage="${stage}", Amt=${totalAmount}`);
    }

    const applied = applyOverridesToRow_(
      {
        Opportunity_UID: uid,
        Account: account,
        Opportunity_Name: oppName,
        Start_Date: startDate,
        End_Date: endDate,
        Currency: currency,
        Total_Amount: totalAmount,
        Capability: capability,
        Tech_Fee_Paying: techFeePaying
      },
      overrides[uid]
    );
    if (applied) debugInfo.push('Override Applied');

    const rate = globalCfg.currencyMap[currency] || 1;
    const totalUsd = totalAmount * rate;

    const monthlyMode = (globalCfg.params.partialMode || 'SIMPLE').toUpperCase();
    const monthly = calculateMonthlyRevenue(totalUsd, startDate, endDate, monthlyMode);
    Object.keys(monthly).forEach(k => monthsSet.add(k));

    masterRows.push({
      Opportunity_UID: uid,
      Account: account,
      Opportunity_Name: oppName,
      Capability: capability || (cfg.amountField === 'Revenue' ? 'Other/Shared' : 'Tech Fee'),
      Start_Date: startDate,
      End_Date: endDate,
      Total_USD: totalUsd,
      Monthly: monthly,
      Currency: currency,
      Tech_Fee_Paying: techFeePaying,
      Override_Applied: applied,
      Debug_Info: debugInfo.join('; ')
    });
  });

  const months = Array.from(monthsSet).sort();
  writeMasterSheet_(masterSh, masterRows, months, cfg.amountField);

  return { rows: masterRows, months };
}

/**
 * Build combined ledger (Revenue + Tech Fee) for downstream dashboards.
 */
function buildCombinedLedger_(estimate, tech) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, 'MASTER_Ledger');

  const months = mergeMonths_(estimate.months, tech.months);
  const header = ['Opportunity_UID', 'Account', 'Opportunity Name', 'Capability', 'Start Date', 'End Date', 'Total_USD', 'Tech_Fee_Paying?'].concat(months);

  const rows = [];
  estimate.rows.forEach(r => {
    rows.push(rowToArray_(r, months));
  });
  tech.rows.forEach(r => {
    rows.push(rowToArray_(r, months));
  });

  sh.clearContents();
  if (rows.length === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    return;
  }
  sh.getRange(1, 1, rows.length + 1, header.length).setValues([header].concat(rows));
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  for (let c = 1; c <= header.length; c++) sh.autoResizeColumn(c);
}

/**
 * Portfolio Health dashboard sheet.
 */
function buildPortfolioHealth_(estimate, tech, cfg) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, 'Portfolio Health');

  // Default to 2025 for now, or make configurable? User asked for 2025/2026 reporting.
  // Let's default to 2025 as the primary view.
  const targetYear = 2025; 
  const accountMap = aggregateAccounts_(estimate, tech, targetYear);
  const rows = [];

  Object.keys(accountMap).forEach(acc => {
    const rec = accountMap[acc];
    const tier = assignTier_(rec.annualRevenue);
    const targetFee = resolveTargetFee_(tier, cfg.pricing);
    const actualFee = rec.actualFeeAnnual;
    const paying = actualFee > 0;
    const status = generateClientActionPlan({
      revenue: rec.annualRevenue,
      targetFee,
      actualFee,
      techFeePaying: paying
    });
    rows.push([
      acc,
      rec.annualRevenue,
      tier,
      targetFee,
      actualFee,
      targetFee - actualFee,
      status.status,
      status.plan
    ]);
  });

  const header = ['Client Name', 'Annual SEO Revenue (USD)', 'Target Tech Tier', 'Target Tech Fee (USD)', 'Actual Tech Fee (USD)', 'Variance', 'Status', 'Recommended Action'];
  sh.clearContents();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.getRange(2, 2, rows.length, 4).setNumberFormat('#,##0');
    sh.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0');
  }
  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(1, rows.length + 1), header.length).createFilter();
  for (let c = 1; c <= header.length; c++) sh.autoResizeColumn(c);
}

/**
 * Renewal Radar sheet.
 */
function buildRenewalRadar_(estimate, tech, cfg) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, 'Renewal Radar');
  const targetYear = 2025;
  const accountMap = aggregateAccounts_(estimate, tech, targetYear);
  const today = cfg.params.today || new Date();
  const lookahead = cfg.params.renewalDays || 90;
  const cutoff = new Date(today.getTime() + lookahead * 24 * 60 * 60 * 1000);

  const rows = [];
  Object.keys(accountMap).forEach(acc => {
    const rec = accountMap[acc];
    if (!rec.earliestEnd || rec.earliestEnd > cutoff) return;
    const tier = assignTier_(rec.annualRevenue);
    const targetFee = resolveTargetFee_(tier, cfg.pricing);
    const actualFee = rec.actualFeeAnnual;
    const paying = actualFee > 0;
    const status = generateClientActionPlan({
      revenue: rec.annualRevenue,
      targetFee,
      actualFee,
      techFeePaying: paying
    });
    if (status.status.indexOf('🟢') === 0) return; // skip healthy
    rows.push([
      acc,
      rec.annualRevenue,
      tier,
      targetFee,
      actualFee,
      status.status,
      status.plan,
      rec.earliestEnd
    ]);
  });

  const header = ['Client Name', 'Annual SEO Revenue (USD)', 'Target Tech Tier', 'Target Tech Fee (USD)', 'Actual Tech Fee (USD)', 'Status', 'Recommended Action', 'Renewal End Date'];
  sh.clearContents();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.getRange(2, 2, rows.length, 3).setNumberFormat('#,##0');
    sh.getRange(2, 8, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  }
  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(1, rows.length + 1), header.length).createFilter();
  for (let c = 1; c <= header.length; c++) sh.autoResizeColumn(c);
}

/*********************
 * Helper functions
 *********************/
function readCurrencyMap_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const ccy = safeStr_(values[i][0]);
    const rate = toNumber_(values[i][1]);
    if (ccy) map[ccy] = rate || 1;
  }
  return map;
}

function readSkuPricing_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const tier = safeStr_(values[i][0]);
    const annual = toNumber_(values[i][3]);
    if (tier) map[tier] = annual || 0;
  }
  return map;
}

function readParams_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const params = { today: new Date(), renewalDays: 90, partialMode: 'SIMPLE' };
  for (let i = 1; i < values.length; i++) {
    const key = safeStr_(values[i][0]);
    const val = values[i][1];
    if (key === 'Today' && val) params.today = parseDate_(val);
    if (key === 'Renewal_Lookahead_Days') params.renewalDays = toNumber_(val) || 90;
    if (key === 'Partial_Month_Mode') params.partialMode = safeStr_(val).toUpperCase() || 'SIMPLE';
  }
  return params;
}

function readOverrides_(sh) {
  const values = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  const map = {};
  values.forEach(r => {
    const uid = safeStr_(r[0]);
    const field = safeStr_(r[1]).toLowerCase();
    const newValue = r[2];
    if (!uid || !field) return;
    map[uid] = map[uid] || {};
    map[uid][field] = newValue;
  });
  return map;
}

function applyOverridesToRow_(row, overrideObj) {
  if (!overrideObj) return false;
  let applied = false;
  Object.keys(overrideObj).forEach(field => {
    const matchKey = Object.keys(row).find(k => safeStr_(k).toLowerCase() === field);
    if (matchKey) {
      row[matchKey] = overrideObj[field];
      applied = true;
    }
  });
  return applied;
}

function generateOpportunityUid_(account, opportunity) {
  const base = `${account || ''}::${opportunity || ''}`;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base);
  const hex = digest.map(b => (b + 256).toString(16).slice(-2)).join('');
  return hex.slice(0, 12);
}

function parseDate_(value) {
  if (value instanceof Date) return value;
  if (value === null || value === undefined || value === '') return null;
  // Handle dd/mm/yyyy or dd-mm-yyyy
  if (typeof value === 'string') {
    const str = value.trim();
    const m = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (m) {
      const day = parseInt(m[1], 10);
      const month = parseInt(m[2], 10) - 1;
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function monthKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

function calculateMonthlyRevenue(amount, startDate, endDate, partialMode) {
  if (!startDate || !endDate || isNaN(startDate) || isNaN(endDate) || amount === 0) return {};
  const mode = (partialMode || 'SIMPLE').toUpperCase();
  if (mode === 'PRORATE') {
    return spreadProratedByDay_(amount, startDate, endDate);
  }
  return spreadEvenly_(amount, startDate, endDate);
}

function spreadEvenly_(amount, startDate, endDate) {
  const months = iterateMonths_(startDate, endDate);
  const monthly = {};
  const perMonth = amount / Math.max(1, months.length);
  months.forEach(m => { monthly[m] = perMonth; });
  return monthly;
}

function spreadProratedByDay_(amount, startDate, endDate) {
  const monthly = {};
  const totalDays = Math.max(1, Math.round((endDate - startDate) / 86400000) + 1);
  const perDay = amount / totalDays;
  let cursor = new Date(startDate);
  while (cursor <= endDate) {
    const mk = monthKey_(cursor);
    monthly[mk] = (monthly[mk] || 0) + perDay;
    cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1);
  }
  return monthly;
}

function iterateMonths_(startDate, endDate) {
  const out = [];
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  let cursor = start;
  while (cursor <= end) {
    out.push(monthKey_(cursor));
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
  }
  return out;
}

function categorizeRevenue(roleString) {
  const role = (roleString || '').toUpperCase();
  if (role.includes('SEO')) return 'SEO Revenue';
  if (role.includes('PUBLIC RELATIONS') || role.includes('PR')) return 'Digital PR Revenue';
  if (role.includes('ASO') || role.includes('APP STORE')) return 'ASO Revenue';
  return 'Other/Shared';
}

function writeMasterSheet_(sheet, rows, months, amountField) {
  const header = ['Opportunity_UID', 'Account', 'Opportunity Name', 'Capability', 'Start Date', 'End Date', 'Total_USD', 'Tech_Fee_Paying?'].concat(months);
  sheet.clearContents();
  if (rows.length === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    return;
  }
  const values = rows.map(r => rowToArray_(r, months));
  sheet.getRange(1, 1, values.length + 1, header.length).setValues([header].concat(values));
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  for (let c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);
}

function rowToArray_(row, months) {
  const base = [
    row.Opportunity_UID,
    row.Account,
    row.Opportunity_Name,
    row.Capability,
    row.Start_Date,
    row.End_Date,
    row.Total_USD,
    row.Tech_Fee_Paying ? 'Yes' : 'No'
  ];
  months.forEach(m => base.push(row.Monthly[m] || 0));
  return base;
}

function mergeMonths_(a, b) {
  const set = new Set();
  (a || []).forEach(m => set.add(m));
  (b || []).forEach(m => set.add(m));
  return Array.from(set).sort();
}

function sumValues_(arr) {
  return (arr || []).reduce((s, v) => s + (Number(v) || 0), 0);
}

function aggregateAccounts_(estimate, tech, targetYear) {
  const map = {};
  const year = targetYear || new Date().getFullYear();
  const yearPrefix = `${year}-`;

  estimate.rows.forEach(r => {
    const acc = r.Account || 'Unknown';
    map[acc] = map[acc] || { revenueMonthly: {}, techMonthly: {}, startDates: [], endDates: [] };
    Object.keys(r.Monthly || {}).forEach(mk => {
      if (mk.startsWith(yearPrefix)) {
        map[acc].revenueMonthly[mk] = (map[acc].revenueMonthly[mk] || 0) + (r.Monthly[mk] || 0);
      }
    });
    if (r.Start_Date) map[acc].startDates.push(r.Start_Date);
    if (r.End_Date) map[acc].endDates.push(r.End_Date);
  });

  tech.rows.forEach(r => {
    const acc = r.Account || 'Unknown';
    map[acc] = map[acc] || { revenueMonthly: {}, techMonthly: {}, startDates: [], endDates: [] };
    Object.keys(r.Monthly || {}).forEach(mk => {
      if (mk.startsWith(yearPrefix)) {
        map[acc].techMonthly[mk] = (map[acc].techMonthly[mk] || 0) + (r.Monthly[mk] || 0);
      }
    });
    if (r.Start_Date) map[acc].startDates.push(r.Start_Date);
    if (r.End_Date) map[acc].endDates.push(r.End_Date);
  });

  Object.keys(map).forEach(acc => {
    const rec = map[acc];
    // Calendar Year Revenue: Sum of all monthly revenue for the target year
    const revenueMonths = Object.keys(rec.revenueMonthly);
    rec.annualRevenue = revenueMonths.reduce((s, m) => s + (rec.revenueMonthly[m] || 0), 0);

    const techMonths = Object.keys(rec.techMonthly);
    rec.actualFeeAnnual = techMonths.reduce((s, m) => s + (rec.techMonthly[m] || 0), 0);

    rec.earliestEnd = rec.endDates.length ? new Date(Math.min.apply(null, rec.endDates.map(d => d.getTime ? d.getTime() : new Date(d).getTime()))) : null;
  });

  return map;
}

function assignTier_(annualRevenue) {
  if (annualRevenue > 500000) return 'Tier A';
  if (annualRevenue >= 200000) return 'Tier B';
  if (annualRevenue >= 75000) return 'Tier C';
  if (annualRevenue >= 0) return 'Tier D';
  return 'Tier D';
}

function resolveTargetFee_(tier, pricingMap) {
  return pricingMap[tier] || 0;
}

function generateClientActionPlan(client) {
  const revenue = client.revenue || 0;
  const targetFee = client.targetFee || 0;
  const actualFee = client.actualFee || 0;
  const techFeePaying = !!client.techFeePaying;

  if (revenue > 200000 && !techFeePaying) {
    return { status: '🔴 CRITICAL MISS', plan: "Client is consuming Enterprise resources. Must add 'Tech Pro' ($12k) at next renewal. Risk of service degradation if not actioned." };
  }
  if (targetFee > actualFee) {
    return { status: '🟠 GAP IDENTIFIED', plan: "Legacy Pricing detected. Propose 'Glide Path': Renewal Year 1 at $3k (50% discount), Year 2 at full price." };
  }
  if (revenue < 50000 && targetFee > revenue * 0.1) {
    return { status: '⚠️ FEE TOO HIGH', plan: "Standard fee exceeds 10% of retainer. Downgrade target to 'Tech Starter' ($2.4k) to preserve deal margin." };
  }
  if (actualFee >= targetFee && targetFee > 0) {
    return { status: '🟢 HEALTHY', plan: 'No action needed. Client fully funds their tier.' };
  }
  return { status: '🟠 GAP IDENTIFIED', plan: 'Tech fee below target. Review SOW and propose glide path.' };
}
