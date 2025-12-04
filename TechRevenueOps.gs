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
  
  // 1. Import Data from Emails
  importOpportunitiesAndAccounts();
  importProjectsAndOppLookup();
  importOppsAndCRs();

  // Clear Debug Sheet at start of pipeline
  const ss = SpreadsheetApp.getActive();
  const debugSh = getOrCreateSheet_(ss, 'RevenueOps_DEBUG');
  debugSh.clear();
  debugSh.appendRow(['Source', 'Stage', 'Message', 'Details']);
  
  // 2. Build Lookups
  const lookups = buildLookups_(ss);

  const estimate = rebuildMasterForSource_(SOURCE_CONFIG.Estimate, cfg, lookups);
  const techFee = rebuildMasterForSource_(SOURCE_CONFIG.TechFee, cfg, lookups);

  buildCombinedLedger_(estimate, techFee);
  buildPortfolioHealth_(estimate, techFee, cfg);
  buildRenewalRadar_(estimate, techFee, cfg);

  return 'Revenue Ops pipeline refreshed (master ledger + dashboards).';
}

function buildLookups_(ss) {
  // A. Project Lookup (Opp Name -> Project Details)
  const projSh = ss.getSheetByName('Projects_RAW_Data_Import');
  const projectLookup = {}; // Opp Name -> { start, end, name, projectId, oppId }
  const projectById = {};   // Project ID -> { name, start, end }
  const projectByOppId = {}; // Opp ID -> { projectId, projectName, start, end }

  if (projSh && projSh.getLastRow() > 1) {
    const data = projSh.getDataRange().getValues();
    const headers = data.shift();
    const hMap = headers.reduce((acc, h, i) => { acc[String(h).trim()] = i; return acc; }, {});
    
    // Expected Headers in Projects_RAW_Data_Import:
    // Project: Project Name, Account, Start Date, End Date, Stage, Opportunity, Project: ID, Opportunity: Opportunity ID
    
    data.forEach(r => {
      const name = safeStr_(r[hMap['Project: Project Name']]);
      const oppName = safeStr_(r[hMap['Opportunity: Opportunity Name']]); // Note: Check exact header from email import
      const oppId = safeStr_(r[hMap['Opportunity: Opportunity ID']]);
      const projId = safeStr_(r[hMap['Project: ID']]);
      const start = parseDate_(r[hMap['Start Date']]);
      const end = parseDate_(r[hMap['End Date']]);

      if (name.toUpperCase().startsWith('OPP_')) return;
      if (name.indexOf('Client Admin & Expense Centre') !== -1) return;

      const account = safeStr_(r[hMap['Account']]); // Capture Account
      const details = { start, end, name, projectId: projId, oppId, account };

      if (oppName) {
        projectLookup[oppName] = projectLookup[oppName] || [];
        projectLookup[oppName].push(details);
      }
      if (projId) projectById[projId] = details;
      if (oppId) {
        projectByOppId[oppId] = projectByOppId[oppId] || [];
        projectByOppId[oppId].push(details);
      }
    });
  }

  // B. Estimate -> Opportunity Map
  const estOppSh = ss.getSheetByName('Estimate_to_Opportunity_Map');
  const estToOppId = {}; // Estimate ID -> Opportunity ID
  if (estOppSh && estOppSh.getLastRow() > 1) {
    const data = estOppSh.getDataRange().getValues();
    // Headers: Account, Opportunity Name, Estimate Name, Estimate ID, Opportunity ID
    // Assuming simple column order or map headers dynamically
    // Let's map dynamically to be safe
    const headers = data.shift();
    const hMap = headers.reduce((acc, h, i) => { acc[String(h).trim()] = i; return acc; }, {});
    
    data.forEach(r => {
      const estId = safeStr_(r[hMap['Estimate: ID']]);
      const oppId = safeStr_(r[hMap['Opportunity: Opportunity ID']]);
      if (estId && oppId) estToOppId[estId] = oppId;
    });
  }

  // D. Estimate Start Dates (Opp ID -> Earliest Start)
  const estSh = ss.getSheetByName('Estimate_RAW_Data_Import');
  const oppStartDates = {}; // Opp ID -> Earliest Start Date
  if (estSh && estSh.getLastRow() > 1) {
    const data = estSh.getDataRange().getValues();
    const headers = data.shift();
    const hMap = headers.reduce((acc, h, i) => { acc[String(h).trim()] = i; return acc; }, {});
    
    // Required: 'Opportunity: Opportunity ID', 'Start Date'
    data.forEach(r => {
      const oppId = safeStr_(r[hMap['Opportunity: Opportunity ID']]);
      const start = parseDate_(r[hMap['Start Date']]);
      if (oppId && start) {
        if (!oppStartDates[oppId] || start < oppStartDates[oppId]) {
          oppStartDates[oppId] = start;
        }
      }
    });
  }

  // C. Opps and CRs (Opp ID -> Parent Opp) + Parent -> Earliest Child Start
  const crSh = ss.getSheetByName('Opps_and_CRs_RAW_Import');
  const oppToParent = {}; // Opp ID -> Parent Opp String (or ID if parsed)
  const parentToChildStart = {}; // Parent Opp ID -> Earliest Child Start Date
  
  if (crSh && crSh.getLastRow() > 1) {
    const data = crSh.getDataRange().getValues();
    const headers = data.shift();
    const hMap = headers.reduce((acc, h, i) => { acc[String(h).trim()] = i; return acc; }, {});
    
    data.forEach(r => {
      const oppId = safeStr_(r[hMap['Opportunity ID']]);
      const parent = safeStr_(r[hMap['Parent Opportunity']]);
      if (oppId && parent) {
        oppToParent[oppId] = parent;
        
        // Extract Parent ID
        const idMatch = parent.match(/\b006[a-zA-Z0-9]{12,15}\b/);
        if (idMatch) {
          const parentId = idMatch[0];
          const childStart = oppStartDates[oppId];
          if (childStart) {
            if (!parentToChildStart[parentId] || childStart < parentToChildStart[parentId]) {
              parentToChildStart[parentId] = childStart;
            }
          }
        }
      }
    });
  }

  return { projectLookup, projectById, projectByOppId, estToOppId, oppToParent, parentToChildStart };
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
    
    // Clear entire sheet to ensure clean slate
    rawSh.clear();
    
    // 1. Write Headers (Force them from Config to ensure "Opportunity" exists)
    rawSh.getRange(1, 1, 1, cfg.requiredHeaders.length).setValues([cfg.requiredHeaders]).setFontWeight('bold');
    
    // 2. Write Data (Skip header row from source, write from Row 2)
    const rows = data.slice(1);
    if (rows.length > 0) {
      // Ensure we don't write more columns than we have headers for, or handle mismatch?
      // User implies data exists in Col F. We'll write all columns available in source rows.
      // Ideally we should slice rows to match header length if source has extra junk, 
      // but for now let's just write what we got.
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
function rebuildMasterForSource_(cfg, globalCfg, lookups) {
  const ss = SpreadsheetApp.getActive();
  const rawSh = ss.getSheetByName(cfg.rawSheet);
  const ovSh = ss.getSheetByName(cfg.overrideSheet);
  const masterSh = ss.getSheetByName(cfg.masterSheet);
  if (!rawSh || !ovSh || !masterSh) throw new Error(`Missing shadow table for ${cfg.rawSheet}`);

  // DEBUG SETUP
  const debugLog = []; 
  const log = (stage, msg, details) => debugLog.push([cfg.rawSheet, stage, msg, typeof details === 'object' ? JSON.stringify(details) : details]);

  const headers = rawSh.getRange(1, 1, 1, rawSh.getLastColumn()).getDisplayValues()[0];
  log('Init', 'Raw Headers Found', headers);

  const rawRows = rawSh.getLastRow() > 1 ? 
    rawSh.getRange(2, 1, rawSh.getLastRow() - 1, headers.length).getValues() : [];
  const overrides = readOverrides_(ovSh);

  const monthsSet = new Set();
  const masterRows = [];
  
  const headerMap = headers.reduce((acc, h, idx) => { 
    acc[String(h).trim().toLowerCase()] = idx; 
    return acc; 
  }, {});
  
  const getIdx = (name) => {
    const key = String(name).trim().toLowerCase();
    const idx = headerMap[key];
    if (idx === undefined) log('HeaderMap', 'Missing Header', key);
    return idx;
  };

  // --- PASS 1: Calculate Scaling Factors ---
  const scalingFactors = {}; 
  if (cfg === SOURCE_CONFIG.Estimate) {
    const fm = cfg.fieldMap;
    const oppTotals = {}; 
    const groupUids = {}; 
      
    rawRows.forEach((raw, i) => {
      const acc = safeStr_(raw[getIdx(fm.account)]);
      const opp = safeStr_(raw[getIdx(fm.opportunity)]);
      const key = `${acc}::${opp}`; 
      const hours = toNumber_(raw[getIdx(fm.hours)]);
      const rate = toNumber_(raw[getIdx(fm.billRate)]);
      const estId = safeStr_(raw[getIdx(fm.estimateId)]);
      
      // Use Estimate ID as group UID if available
      if (estId) groupUids[key] = estId;
        
      if (!oppTotals[key]) oppTotals[key] = 0;
      oppTotals[key] += (hours * rate);
    });

    Object.keys(oppTotals).forEach(key => {
      const parts = key.split('::');
      const groupUid = groupUids[key] || generateOpportunityUid_(parts[0], parts[1]); 
      const rawTotal = oppTotals[key];
      const ov = overrides[groupUid];
      if (ov) {
         const targetVal = ov['total_usd'] || ov['target_revenue'];
         if (targetVal) {
           const target = Number(targetVal);
           if (!isNaN(target) && rawTotal > 0) {
             scalingFactors[key] = target / rawTotal; 
           }
         }
      }
    });
  }

  // --- PASS 2: Generate Rows ---
  rawRows.forEach((raw, i) => {
    const get = name => raw[getIdx(name)];
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
      
    let role = '';
    let region = '';
    let hours = 0;
    let billRate = 0;
    
    // New Lookup Fields
    let oppId = '';
    let projId = '';
    let projName = '';

    if (cfg === SOURCE_CONFIG.Estimate) {
      const fm = cfg.fieldMap;
      account = safeStr_(get(fm.account));
      oppName = safeStr_(get(fm.opportunity));
      const estId = safeStr_(get(fm.estimateId));
      uid = estId || generateOpportunityUid_(account, oppName); // Use Estimate ID as UID
        
      role = safeStr_(get(fm.role));
      region = safeStr_(get(fm.region));
      hours = toNumber_(get(fm.hours));
      billRate = toNumber_(get(fm.billRate));

      // Resolve Project Details
      const details = resolveProjectDetails_(estId, oppName, lookups, account);
      if (details) {
        oppId = details.oppId;
        projId = details.projectId;
        projName = details.projectName;
        if (details.start) startDate = details.start;
        if (details.end) endDate = details.end;
        debugInfo.push(`Project Linked: ${projName}`);
      }

      // --- WATERFALL DATE LOGIC ---
      // Case A: This is a Change Request (CR)
      if (lookups.oppToParent[oppId]) {
        // CRs must ignore Project dates and use Estimate dates
        startDate = parseDate_(get(fm.startDate));
        endDate = parseDate_(get(fm.endDate));
        debugInfo.push('CR Detected: Using Estimate Dates');
      }
      // Case B: This is a Parent Opportunity of a CR
      else if (lookups.parentToChildStart[oppId]) {
        // Parent ends 1 day before the earliest CR starts
        const childStart = lookups.parentToChildStart[oppId];
        const cutoff = new Date(childStart);
        cutoff.setDate(cutoff.getDate() - 1);
        
        // Use Project Start (or fallback), but cap End Date
        if (!startDate) startDate = parseDate_(get(fm.startDate));
        endDate = cutoff;
        debugInfo.push(`Parent of CR: End Date Capped at ${monthKey_(endDate)}`);
      }
      // Case C: Standard Project (Fallthrough)
      
      if (!startDate) {
        startDate = parseDate_(get(fm.startDate));
        endDate = parseDate_(get(fm.endDate));
      }
        
      currency = safeStr_(get(fm.currency)) || 'USD';
      let rawLineAmount = hours * billRate;
      const groupKey = `${account}::${oppName}`;
      if (scalingFactors[groupKey]) {
        const factor = scalingFactors[groupKey];
        rawLineAmount = rawLineAmount * factor;
        debugInfo.push(`Scaled by ${(factor*100).toFixed(1)}%`);
      }
      totalAmount = rawLineAmount;
      capability = categorizeRevenue(get(fm.role));
      techFeePaying = false;

    } else {
      // Tech Fee
      const fm = cfg.fieldMap;
      uid = safeStr_(get(fm.opportunityUid)) || generateOpportunityUid_(safeStr_(get(fm.account)), safeStr_(get(fm.opportunity)));
      account = safeStr_(get(fm.account));
      oppName = safeStr_(get(fm.opportunity));
      oppId = safeStr_(get(fm.opportunityUid));

      // Try to find project by Opp Name or ID
      if (lookups.projectByOppId[oppId]) {
        const p = lookups.projectByOppId[oppId];
        projId = p.projectId;
        projName = p.name;
        startDate = p.start;
        endDate = p.end;
      } else if (lookups.projectLookup[oppName]) {
        const p = lookups.projectLookup[oppName];
        projId = p.projectId;
        projName = p.name;
        startDate = p.start;
        endDate = p.end;
      }

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
           }
        }
      }
      currency = safeStr_(get(fm.currency)) || 'USD';
      totalAmount = parseCurrency_(get(fm.productAmount));
      capability = safeStr_(get(fm.productName)) || 'Tech Fee';
      techFeePaying = true;
    }

    // 1. Apply overrides
    const rowInputs = {
      Opportunity_UID: uid,
      Account: account,
      Opportunity_Name: oppName,
      Start_Date: startDate,
      End_Date: endDate,
      Currency: currency,
      Total_Amount: totalAmount,
      Capability: capability,
      Tech_Fee_Paying: techFeePaying,
      Resource_Role: role,
      Resource_Region: region,
      Hours: hours,
      Bill_Rate: billRate,
      Opportunity_ID: oppId,
      Project_ID: projId,
      Project_Name: projName
    };
    
    const inputsApplied = applyOverridesToRow_(rowInputs, overrides[uid]);
    
    // Update locals
    uid = safeStr_(rowInputs.Opportunity_UID);
    account = safeStr_(rowInputs.Account);
    oppName = safeStr_(rowInputs.Opportunity_Name);
    startDate = rowInputs.Start_Date;
    endDate = rowInputs.End_Date;
    currency = safeStr_(rowInputs.Currency);
    totalAmount = toNumber_(rowInputs.Total_Amount);
    capability = safeStr_(rowInputs.Capability);
    techFeePaying = !!rowInputs.Tech_Fee_Paying;
    role = safeStr_(rowInputs.Resource_Role);
    region = safeStr_(rowInputs.Resource_Region);
    hours = toNumber_(rowInputs.Hours);
    billRate = toNumber_(rowInputs.Bill_Rate);
    oppId = safeStr_(rowInputs.Opportunity_ID);
    projId = safeStr_(rowInputs.Project_ID);
    projName = safeStr_(rowInputs.Project_Name);

    // 2. Calculate derived USD
    const rate = globalCfg.currencyMap[currency] || 1;
    let totalUsd = totalAmount * rate;

    // 3. Apply overrides to outputs
    let outputsApplied = false;
    if (cfg !== SOURCE_CONFIG.Estimate) {
      const rowOutputs = { Total_USD: totalUsd };
      outputsApplied = applyOverridesToRow_(rowOutputs, overrides[uid]);
      totalUsd = toNumber_(rowOutputs.Total_USD);
    }

    const applied = inputsApplied || outputsApplied;
    const monthlyMode = (globalCfg.params.partialMode || 'SIMPLE').toUpperCase();
    const monthly = calculateMonthlyRevenue(totalUsd, startDate, endDate, monthlyMode);
    Object.keys(monthly).forEach(k => monthsSet.add(k));

    masterRows.push({
      Opportunity_UID: uid, // This is Estimate ID for Estimates
      Account: account,
      Opportunity_Name: oppName,
      Capability: capability || (cfg.amountField === 'Revenue' ? 'Other/Shared' : 'Tech Fee'),
      Resource_Role: role,
      Resource_Region: region,
      Hours: hours,
      Bill_Rate: billRate,
      Start_Date: startDate,
      End_Date: endDate,
      Total_USD: totalUsd,
      Monthly: monthly,
      Currency: currency,
      Tech_Fee_Paying: techFeePaying,
      Override_Applied: applied,
      Debug_Info: debugInfo.join('; '),
      Opportunity_ID: oppId, // NEW
      Project_ID: projId,    // NEW
      Project_Name: projName // NEW
    });
  });

  const months = Array.from(monthsSet).sort();
  writeMasterSheet_(masterSh, masterRows, months, cfg.amountField);
  
  // Write Debug Log
  const debugSh = getOrCreateSheet_(ss, 'RevenueOps_DEBUG');
  if (debugLog.length > 0) {
    debugSh.getRange(debugSh.getLastRow() + 1, 1, debugLog.length, 4).setValues(debugLog);
  }

  return { rows: masterRows, months };
}

function resolveProjectDetails_(estId, oppName, lookups, account) {
  // Helper to find best match in a list of projects
  const findBestMatch = (projects) => {
    if (!projects || projects.length === 0) return null;
    if (projects.length === 1) return projects[0]; // If only one, assume it's correct (or check account?)
    
    // If multiple, filter by account
    if (account) {
      const match = projects.find(p => safeStr_(p.account).toLowerCase() === safeStr_(account).toLowerCase());
      if (match) return match;
      // Fuzzy match? e.g. "Converse" vs "Converse EMEA"
      const fuzzy = projects.find(p => safeStr_(p.account).toLowerCase().includes(safeStr_(account).toLowerCase()) || safeStr_(account).toLowerCase().includes(safeStr_(p.account).toLowerCase()));
      if (fuzzy) return fuzzy;
    }
    return projects[0]; // Fallback to first
  };

  // 1. Try Estimate ID -> Opportunity ID
  let oppId = lookups.estToOppId[estId];
  
  // 2. If we have Opp ID, check if it has a Parent Opportunity (CR logic)
  if (oppId && lookups.oppToParent[oppId]) {
    const parentStr = lookups.oppToParent[oppId];
    // Parent might be "Name - ID - Name" or just ID. 
    // Let's try to extract ID if present (regex for 15/18 char ID starting with 006)
    const idMatch = parentStr.match(/\b006[a-zA-Z0-9]{12,15}\b/);
    if (idMatch) {
      const parentId = idMatch[0];
      if (lookups.projectByOppId[parentId]) {
        const p = findBestMatch(lookups.projectByOppId[parentId]);
        if (p) return { projectId: p.projectId, projectName: p.name, oppId: oppId, start: p.start, end: p.end };
      }
    } else {
       // Try lookup by Parent Name (if parentStr is a name)
       if (lookups.projectLookup[parentStr]) {
         const p = findBestMatch(lookups.projectLookup[parentStr]);
         if (p) return { projectId: p.projectId, projectName: p.name, oppId: oppId, start: p.start, end: p.end };
       }
    }
  }

  // 3. If we have Opp ID, check direct Project link
  if (oppId && lookups.projectByOppId[oppId]) {
    const p = findBestMatch(lookups.projectByOppId[oppId]);
    if (p) return { projectId: p.projectId, projectName: p.name, oppId: oppId, start: p.start, end: p.end };
  }

  // 4. Fallback: Lookup by Opportunity Name
  if (lookups.projectLookup[oppName]) {
    const p = findBestMatch(lookups.projectLookup[oppName]);
    if (p) return { projectId: p.projectId, projectName: p.name, oppId: p.oppId || oppId, start: p.start, end: p.end };
  }

  return { projectId: '', projectName: '', oppId: oppId || '' };
}

/**
 * Build combined ledger (Revenue + Tech Fee) for downstream dashboards.
 */
function buildCombinedLedger_(estimate, tech) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, 'MASTER_Ledger');
  const debugSh = ss.getSheetByName('RevenueOps_DEBUG');

  const months = mergeMonths_(estimate.months, tech.months);
  const header = ['Opportunity_UID', 'Account', 'Opportunity Name', 'Capability', 'Resource Role', 'Resource Region', 'Hours', 'Bill Rate', 'Start Date', 'End Date', 'Total_USD', 'Tech_Fee_Paying?', 'Opportunity_ID', 'Project_ID', 'Project_Name'].concat(months);

  const rows = [];
  estimate.rows.forEach(r => {
    if (debugSh && r.Opportunity_Name && r.Opportunity_Name.includes('Swarovski')) {
       debugSh.appendRow(['CombinedLedger', 'Loop', 'Row Check', JSON.stringify({ role: r.Resource_Role, region: r.Resource_Region })]);
    }
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
  // Updated Header Layout
  const header = [
    amountField === 'Revenue' ? 'Estimate ID' : 'Opportunity_UID', 
    'Account', 
    'Opportunity Name', 
    'Capability', 
    'Resource Role',    
    'Resource Region',  
    'Hours',            
    'Bill Rate',        
    'Start Date', 
    'End Date', 
    'Total_USD', 
    'Tech_Fee_Paying?',
    'Opportunity_ID', // NEW
    'Project_ID',     // NEW
    'Project_Name'    // NEW
  ].concat(months);

  sheet.clearContents();
  if (rows.length === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    return;
  }
    
  const values = rows.map(r => {
    const base = [
      r.Opportunity_UID,
      r.Account,
      r.Opportunity_Name,
      r.Capability,
      r.Resource_Role,    
      r.Resource_Region,  
      r.Hours,            
      r.Bill_Rate,        
      r.Start_Date,
      r.End_Date,
      r.Total_USD,
      r.Tech_Fee_Paying ? 'Yes' : 'No',
      r.Opportunity_ID, // NEW
      r.Project_ID,     // NEW
      r.Project_Name    // NEW
    ];
    months.forEach(m => base.push(r.Monthly[m] || 0));
    return base;
  });

  sheet.getRange(1, 1, values.length + 1, header.length).setValues([header].concat(values));
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function rowToArray_(row, months) {
  const base = [
    row.Opportunity_UID,
    row.Account,
    row.Opportunity_Name,
    row.Capability,
    row.Resource_Role,
    row.Resource_Region,
    row.Hours,
    row.Bill_Rate,
    row.Start_Date,
    row.End_Date,
    row.Total_USD,
    row.Tech_Fee_Paying ? 'Yes' : 'No',
    row.Opportunity_ID,
    row.Project_ID,
    row.Project_Name
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

function parseCurrency_(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  const str = String(value);
  // Remove everything that is not a digit, dot, or minus
  const clean = str.replace(/[^0-9.-]+/g, '');
  const n = parseFloat(clean);
  return isFinite(n) ? n : 0;
}
