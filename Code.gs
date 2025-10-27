/*************************
 * MENU
 *************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tech Fee Tools')
    .addItem('Build Tech Fee Join', 'Build_Tech_Fee_Join')      // restored
    .addItem('Refresh Fair-Usage Table', 'Build_FairUsage_ForYear')
    .addItem('Create/Update Setup Tab', 'EnsureSetupTab_')
    .addToUi();
}

/*************************
 * 1) REVENUE vs TECH FEE JOIN (restored)
 *************************/
function Build_Tech_Fee_Join() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // Tab names in your file
  const SEO_SHEET  = 'SEO Client Revenue';
  const TOOL_SHEET = 'Tool Revenue';

  // Structures (as per your sheet)
  const SEO_START_ROW = 2;                        // headers row 1
  const SEO_YEAR_COLS = { 2024:3, 2025:4, 2026:5 };
  const TOOL_START_ROW = 5;                       // headers row 4
  const TOOL_YEAR_COLS = { 2024:2, 2025:3, 2026:4 };

  // Ask year
  const resp = ui.prompt('Choose Year', 'Enter 2024, 2025, or 2026 (default 2025):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const year = parseInt((resp.getResponseText() || '2025').trim(), 10);
  if (![2024, 2025, 2026].includes(year)) { ui.alert('Invalid year.'); return; }

  // Get sheets
  const seoSh  = ss.getSheetByName(SEO_SHEET);
  const toolSh = ss.getSheetByName(TOOL_SHEET);
  if (!seoSh) throw new Error(`Missing sheet: ${SEO_SHEET}`);
  if (!toolSh) throw new Error(`Missing sheet: ${TOOL_SHEET}`);

  // Read SEO Client Revenue
  const seoYearCol = SEO_YEAR_COLS[year];
  const seoLastRow = getLastRow_(seoSh, 1, SEO_START_ROW);
  if (seoLastRow < SEO_START_ROW) throw new Error(`No data in ${SEO_SHEET}.`);
  const seoRange = seoSh.getRange(SEO_START_ROW, 1, seoLastRow - SEO_START_ROW + 1, Math.max(2, seoYearCol));
  const seoRaw = seoRange.getValues();
  const seoRows = seoRaw.map(r => ({
    account: safeStr_(r[0]),
    market:  safeStr_(r[1]),
    seoRevenue: toNumber_(r[seoYearCol - 1])
  })).filter(r => r.account);

  // Read Tool Revenue
  const toolYearCol = TOOL_YEAR_COLS[year];
  const toolLastRow = getLastRow_(toolSh, 1, TOOL_START_ROW);
  let techByAccount = new Map();
  if (toolLastRow >= TOOL_START_ROW) {
    const toolRange = toolSh.getRange(TOOL_START_ROW, 1, toolLastRow - TOOL_START_ROW + 1, Math.max(1, toolYearCol));
    const toolRaw = toolRange.getValues();
    techByAccount = new Map(toolRaw.map(r => [safeStr_(r[0]).toLowerCase(), toNumber_(r[toolYearCol - 1]) || 0]));
  }

  // Output
  const outName = `Revenue vs Tech Fee - ${year}`;
  const outSh = getOrCreateSheet_(ss, outName);
  const header = ['Account','Market','Year','SEO Revenue','Tech Fee Revenue','Paying Tech Fee?'];
  const out = [header];

  seoRows.forEach(r => {
    const tech = techByAccount.get(r.account.toLowerCase()) ?? 0;
    out.push([r.account, r.market, year, r.seoRevenue, tech, tech > 0 ? 'Yes' : 'No']);
  });

  outSh.clearContents();
  outSh.getRange(1, 1, out.length, header.length).setValues(out);

  // Formatting
  outSh.getRange(1,1,1,header.length).setFontWeight('bold');
  if (out.length > 1) outSh.getRange(2,4,out.length-1,2).setNumberFormat('#,##0.00');
  if (outSh.getFilter()) outSh.getFilter().remove();
  outSh.getRange(1,1,out.length,header.length).createFilter();
  outSh.setFrozenRows(1);
  for (let c=1;c<=header.length;c++) outSh.autoResizeColumn(c);

  // use local ss variable (consistent) instead of SpreadsheetApp.getActiveSpreadsheet()
  ss.setActiveSheet(outSh);
  ui.alert(`Done. Wrote ${out.length - 1} rows to "${outName}".`);
}

/*************************
 * 2) FAIR-USAGE BUILDER (uses Setup tab)
 *************************/
function Build_FairUsage_ForYear() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const SEO_SHEET  = 'SEO Client Revenue';
  const TOOL_SHEET = 'Tool Revenue';
  const SEO_START_ROW = 2;
  const SEO_YEAR_COLS = { 2024:3, 2025:4, 2026:5 };
  const TOOL_START_ROW = 5;
  const TOOL_YEAR_COLS = { 2024:2, 2025:3, 2026:4 };

  const resp = ui.prompt('Choose Year', 'Enter 2024, 2025, or 2026 (default 2025):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const year = parseInt((resp.getResponseText() || '2025').trim(), 10);
  if (![2024,2025,2026].includes(year)) { ui.alert('Invalid year.'); return; }

  const seoSh  = ss.getSheetByName(SEO_SHEET);
  const toolSh = ss.getSheetByName(TOOL_SHEET);
  if (!seoSh) throw new Error(`Missing sheet: ${SEO_SHEET}`);
  if (!toolSh) throw new Error(`Missing sheet: ${TOOL_SHEET}`);

  // Load Setup (this also CREATES it if missing)
  const cfg = EnsureSetupTab_();

  // SEO
  const seoYearCol = SEO_YEAR_COLS[year];
  const seoLastRow = getLastRow_(seoSh, 1, SEO_START_ROW);
  const seoRange = seoSh.getRange(SEO_START_ROW, 1, Math.max(0, seoLastRow-SEO_START_ROW+1), Math.max(2, seoYearCol));
  const seoRaw = seoRange.getValues();
  const seoRows = seoRaw.map(r => ({
    account: safeStr_(r[0]),
    market:  safeStr_(r[1]),
    revenue: toNumber_(r[seoYearCol-1])
  })).filter(r => r.account);

  // Tech fee
  const toolYearCol = TOOL_YEAR_COLS[year];
  const toolLastRow = getLastRow_(toolSh, 1, TOOL_START_ROW);
  const toolRange = toolSh.getRange(TOOL_START_ROW, 1, Math.max(0, toolLastRow-TOOL_START_ROW+1), Math.max(1, toolYearCol));
  const toolRaw = toolRange.getValues();
  const techByAccount = new Map();
  toolRaw.forEach(r => {
    const acc = safeStr_(r[0]);
    const tf  = toNumber_(r[toolYearCol-1]);
    if (acc) techByAccount.set(acc.toLowerCase(), tf || 0);
  });

  // --- Tiering + region (mark inactive if no revenue)
  const rows = seoRows.map(r => {
    const revenue = Number(r.revenue) || 0;
    const inactive = revenue <= 0;

    // still look up tech fee, but inactive accounts won't get allocations
    const techFee = techByAccount.get(r.account.toLowerCase()) || 0;
    const paying  = techFee > 0;

    // If inactive, force to "Inactive" with zero base/ceiling
    const tierObj = inactive ? { name:'Inactive', base:0, ceiling:0 } : inferTier_(revenue, cfg.tiers);

    const bandName = cfg.marketBands[r.market] || 'Low';
    const regionMult = cfg.regionalBands[bandName] != null ? cfg.regionalBands[bandName] : 1.00;

    return {
      account: r.account,
      market:  r.market,
      revenue,
      techFee,
      paying,
      inactive,                 // <— new flag
      tier: tierObj.name,
      tierBase: tierObj.base,
      tierCeiling: tierObj.ceiling,
      regionBandName: bandName,
      regionMult
    };
  });


  // --- AccuRanker allocation (skip inactive)
  // OnCrawl allocation using weight-based Base + Contributor pool (policy)
  const totalTokens = Math.floor(cfg.accuCapacity);                   // tokens/month (OnCrawl)
  const pot1 = Math.floor(totalTokens * cfg.basePct);                 // Base pot (65%)
  const pot2 = Math.floor(totalTokens * cfg.poolPct);                 // Contributor pool (25%)
  const buffer = Math.floor(totalTokens * cfg.bufferPct);            // Reserved (~10%)

  // Tier weights (per policy)
  const weights = { 'Tier A': 5, 'Tier B': 3, 'Tier C': 1, 'Tier D': 0.5 };

  // Count active accounts per tier (base applies to active accounts)
  const active = rows.filter(r => !r.inactive);
  const counts = { 'Tier A':0, 'Tier B':0, 'Tier C':0, 'Tier D':0 };
  active.forEach(r => { if (counts[r.tier] != null) counts[r.tier]++; });

  const totalWeight = Object.keys(counts).reduce((s,t) => s + (weights[t] || 0) * counts[t], 0);

  // Weighted tech-fee total among payers for contributor pool (region bump included)
  const payers = rows.filter(r => !r.inactive && r.paying);
  const weightedTechTotal = payers.reduce((s, r) => s + (r.techFee || 0) * (r.regionMult || 1), 0);

  // Assign Base + Contributor; inactive get zeros
  rows.forEach(r => {
    if (r.inactive) {
      r.accuBase = 0;
      r.accuContributor = 0;
      r.accuTotal = 0;
      r.poolSharePct = 0;
      return;
    }

    // Base per account = pot1 * (W_tier / Σ(W_tier × #accounts_in_tier))
    const w = weights[r.tier] || 0;
    r.accuBase = totalWeight > 0 ? Math.floor(pot1 * (w / totalWeight)) : 0;

    // Contributor pool share: normalized by weighted tech fees (techFee * regionMult)
    const weighted = r.paying ? ((r.techFee || 0) * (r.regionMult || 1)) : 0;
    const share = weightedTechTotal > 0 ? (weighted / weightedTechTotal) : 0;
    const extra = r.paying ? Math.floor(pot2 * share) : 0;

    // Cap by tier ceiling (headroom from Base)
    const headroom = Math.max(0, (r.tierCeiling || 0) - r.accuBase);
    r.accuContributor = Math.min(extra, headroom);
    r.accuTotal = r.accuBase + r.accuContributor;
    r.poolSharePct = r.paying ? (share * 100) : 0;
  });

  // --- Enforce total ≤ totalTokens - buffer by reducing contributor first, then base
  (function enforcePoolLimit() {
    const available = Math.max(0, totalTokens - buffer);
    const activeRows = rows.filter(r => !r.inactive);
    let totalAllocated = activeRows.reduce((s,r) => s + (r.accuTotal || 0), 0);
    if (totalAllocated <= available) return;

    // 1) Try reducing contributor pool proportionally
    let totalContributor = activeRows.reduce((s,r) => s + (r.accuContributor || 0), 0);
    if (totalContributor > 0) {
      const neededReduction = Math.min(totalAllocated - available, totalContributor);
      const contributorScale = Math.max(0, (totalContributor - neededReduction) / totalContributor);
      activeRows.forEach(r => {
        r.accuContributor = Math.floor((r.accuContributor || 0) * contributorScale);
        // recompute total but still respect ceiling (accuBase unchanged here)
        r.accuTotal = (r.accuBase || 0) + r.accuContributor;
      });
      totalAllocated = activeRows.reduce((s,r) => s + (r.accuTotal || 0), 0);
      if (totalAllocated <= available) return;
    }

    // 2) Still over: reduce base proportionally (this may reduce cadence but keeps contributor reductions)
    let totalBase = activeRows.reduce((s,r) => s + (r.accuBase || 0), 0);
    if (totalBase > 0) {
      const remainingExcess = Math.max(0, totalAllocated - available);
      const baseScale = Math.max(0, (totalBase - remainingExcess) / totalBase);
      activeRows.forEach(r => {
        r.accuBase = Math.floor((r.accuBase || 0) * baseScale);
        // Ensure we don't exceed tier ceiling after base reduction (shouldn't), recompute total
        const headroom = Math.max(0, (r.tierCeiling || 0) - r.accuBase);
        r.accuContributor = Math.min(r.accuContributor || 0, headroom);
        r.accuTotal = r.accuBase + r.accuContributor;
      });
      totalAllocated = activeRows.reduce((s,r) => s + (r.accuTotal || 0), 0);
    }

    // If still over (extremely unlikely), do a final proportional floor across total (last-resort)
    if (totalAllocated > available && totalAllocated > 0) {
      const finalScale = available / totalAllocated;
      activeRows.forEach(r => {
        const newTotal = Math.floor((r.accuTotal || 0) * finalScale);
        // preserve contributor/base ratio where possible
        const contribRatio = (r.accuContributor || 0) / Math.max(1, r.accuTotal || 1);
        r.accuContributor = Math.floor(newTotal * contribRatio);
        r.accuBase = newTotal - r.accuContributor;
        // enforce ceiling
        const headroom = Math.max(0, (r.tierCeiling || 0) - r.accuBase);
        r.accuContributor = Math.min(r.accuContributor, headroom);
        r.accuTotal = r.accuBase + r.accuContributor;
      });
    }
  })();


  // --- Semrush caps (inactive = 0)
  rows.forEach(r => {
    if (r.inactive) { r.semrushCap = 0; return; }
    const caps = cfg.semrushCaps[r.tier] || cfg.semrushCaps['Default'] || { nonpaying: 50, paying: 100 };
    // use explicit fallbacks for each side and fix typo
    r.semrushCap = r.paying ? (caps.paying ?? 100) : (caps.nonpaying ?? 50);
  });


  // --- OnCrawl cadence (inactive = None)
  rows.forEach(r => {
    r.oncrawlCadence = r.inactive ? 'None' : cadenceFor_(r.tier, r.paying, cfg.crawlCadence);
  });

  // --- OnCrawl starter caps (inactive = 0)
  rows.forEach(r => {
    if (r.inactive) {
      r.oncrawlBase = 0;
    } else {
      const caps = cfg.oncrawlCaps[r.tier] || cfg.oncrawlCaps['Tier D'] || { nonpaying: 2500, paying: 4000 };
      r.oncrawlBase = r.paying ? (caps.paying ?? 4000) : (caps.nonpaying ?? 2500);
    }
  });

  // Output
  const outName = `Tech Fair-Usage — ${year}`;
  const outSh = getOrCreateSheet_(ss, outName);
  const header = [
    'Account','Market','Year',
    'Tier','Pays Tech Fee?','Revenue','Tech Fee',
    'Regional Band','Tech-Fee Share % (Pool)',
    'AccuRanker Base','AccuRanker Contributor','AccuRanker Total',
    'OnCrawl Base','OnCrawl Contributor','OnCrawl Total',
    'Semrush Keyword Cap','OnCrawl Cadence'
  ];
  const out = [header];
  rows.forEach(r => out.push([
    r.account, r.market, year,
    r.tier, r.paying ? 'Yes':'No', r.revenue, r.techFee,
    r.regionBandName, round2_(r.poolSharePct),
    r.accuBase, r.accuContributor, r.accuTotal,
    // OnCrawl uses starter caps for now. Contributor logic can be added later.
    r.oncrawlBase, 0, r.oncrawlBase,
    r.semrushCap, r.oncrawlCadence
  ]));

  outSh.clearContents();
  outSh.getRange(1,1,out.length,header.length).setValues(out);

  // Formatting
  outSh.getRange(1,1,1,header.length).setFontWeight('bold');
  if (out.length > 1) {
    outSh.getRange(2,6,out.length-1,2).setNumberFormat('#,##0');      // revenue, tech fee
    outSh.getRange(2,9,out.length-1,1).setNumberFormat('0.00"%"');    // share %
    outSh.getRange(2,10,out.length-1,3).setNumberFormat('#,##0');     // AccuRanker numbers
    outSh.getRange(2,13,out.length-1,3).setNumberFormat('#,##0');     // OnCrawl numbers
    outSh.getRange(2,16,out.length-1,1).setNumberFormat('#,##0');     // Semrush cap
  }
  if (outSh.getFilter()) outSh.getFilter().remove();
  outSh.getRange(1,1,out.length,header.length).createFilter();
  outSh.setFrozenRows(1);
  for (let c=1;c<=header.length;c++) outSh.autoResizeColumn(c);

  // use local ss variable (consistent) instead of SpreadsheetApp.getActiveSpreadsheet()
  ss.setActiveSheet(outSh);
  ui.alert(`Built "${outName}" with ${rows.length} accounts.`);
}

/*************************
 * 3) SETUP TAB (creates or updates, with 4 cols padding)
 *************************/
function EnsureSetupTab_() {
  const ss = SpreadsheetApp.getActive();
  const shName = 'Setup';
  let sh = ss.getSheetByName(shName);
  if (!sh) sh = ss.insertSheet(shName);

  // If mostly empty, write defaults (ensure exactly 4 columns per row)
  if (sh.getLastRow() < 5) {
    sh.clear();

    const rows = [
      ['Key','Value','Notes',''],
      ['ACCURANKER_CAPACITY',100000,'AccuRanker capacity (≈100k tracking slots).',''],
      ['SPLIT_BASE_PCT',0.65,'~65% capacity as Base by Tier.',''],
      ['SPLIT_POOL_PCT',0.25,'~25% Contributor Pool (paying accounts only).',''],
      ['SPLIT_BUFFER_PCT',0.10,'~10% buffer for launches/incidents.',''],
      ['','','',''],
      ['Revenue Tiers','Min','Max','Base / Ceiling'],
      // These Base / Ceiling values are for AccuRanker (base = default non-paying allocation; ceiling = max)
      ['Tier A',500000,'', '800 / 2000'],
      ['Tier B',200000,499999,'500 / 1200'],
      ['Tier C',50000,199999,'250 / 600'],
      ['Tier D',0,49999,'100 / 250'],
      ['','','',''],
      ['Regional Bands','Multiplier','','Set Market→Band below'],
      ['Top',1.20,'',''],
      ['Mid',1.10,'',''],
      ['Low',1.00,'',''],
      ['','','',''],
      ['Semrush Caps','Non-paying','Paying','(per client)'],
      ['Tier A',200,400,''],
      ['Tier B',150,300,''],
      ['Tier C',100,200,''],
      ['Tier D',50,100,''],
      ['','','',''],
      ['Crawl Cadence Rules','Non-paying','Paying','(OnCrawl)'],
      ['Tier A','Monthly','Weekly/Fortnightly',''],
      ['Tier B','Bi-monthly or Quarterly','Monthly',''],
      ['Tier C','Quarterly','Quarterly',''],
      ['Tier D','One-off / Quarterly by request','One-off / Quarterly by request',''],
      ['','','',''],
      ['OnCrawl Starter Caps','Non-paying','Paying','(monthly starter defaults)'],
      ['Tier A',25000,40000,''],
      ['Tier B',10000,18000,''],
      ['Tier C',5000,8000,''],
      ['Tier D',2500,4000,''],
      ['','','',''],
      ['Market→Regional Band','Band','','Set one band per Market'],
      ['UK','Top','',''],
      ['ZA','Top','',''],
      ['US','Mid','',''],
      ['FR','Low','','']
    ];

    // Guarantee 4 columns on every row
    const rows4 = rows.map(r => [r[0] ?? '', r[1] ?? '', r[2] ?? '', r[3] ?? '']);

    sh.getRange(1,1,rows4.length,4).setValues(rows4);
    sh.setFrozenRows(1);
    for (let c=1;c<=4;c++) sh.autoResizeColumn(c);
  }

  // Read config
  const values = sh.getDataRange().getDisplayValues();

  const val = (key, def) => {
    const row = values.find(r => r[0] === key);
    return row ? row[1] : def;
  };

  // Allow a typo safety net from previous version
  const fallbackCap = parseInt(val('ACCUNOTREAL', '100000'), 10);
  const capacity = parseInt(val('ACCURANKER_CAPACITY', fallbackCap || 100000), 10);
  const basePct   = parseFloat(val('SPLIT_BASE_PCT', 0.65));
  const poolPct   = parseFloat(val('SPLIT_POOL_PCT', 0.25));
  const bufferPct = parseFloat(val('SPLIT_BUFFER_PCT', 0.10));

  // Parse “Revenue Tiers” block with “Base / Ceiling” column
  const tiers = [];
  readTable_(values, 'Revenue Tiers', row => {
    const name = row[0];
    if (/^Tier\s+[A-D]$/i.test(name)) {
      const baseCeil = String(row[3] || '').split('/');
      const base = parseInt((baseCeil[0] || '').trim(), 10);
      const ceiling = parseInt((baseCeil[1] || '').trim(), 10);
      tiers.push({
        name,
        min: toNumber_(row[1]),
        max: row[2] ? toNumber_(row[2]) : Number.POSITIVE_INFINITY,
        base: isFinite(base) ? base : 0,
        ceiling: isFinite(ceiling) ? ceiling : 0
      });
    }
  });

  // Regional bands
  const regionalBands = {};
  readTable_(values, 'Regional Bands', row => {
    if (row[0] && row[1]) regionalBands[row[0]] = parseFloat(row[1]);
  });

  // Semrush caps
  const semrushCaps = {};
  readTable_(values, 'Semrush Caps', row => {
    if (row[0] && row[1] && row[2]) {
      semrushCaps[row[0]] = { nonpaying: parseInt(row[1],10), paying: parseInt(row[2],10) };
    }
  });

  // Crawl cadence
  const crawlCadence = {};
  readTable_(values, 'Crawl Cadence Rules', row => {
    if (row[0] && row[1] && row[2]) {
      crawlCadence[row[0]] = { nonpaying: row[1], paying: row[2] };
    }
  });

  // Market→Band mapping
  const marketBands = {};
  readTable_(values, 'Market→Regional Band', row => {
    if (row[0] && row[1]) marketBands[row[0]] = row[1];
  });

  // OnCrawl starter caps
  const oncrawlCaps = {};
  readTable_(values, 'OnCrawl Starter Caps', row => {
    if (row[0] && /Tier/.test(row[0]) && row[1] && row[2]) {
      oncrawlCaps[row[0]] = { nonpaying: parseInt(row[1], 10), paying: parseInt(row[2], 10) };
    }
  });

  return {
    accuCapacity: capacity,
    basePct, poolPct, bufferPct,
    tiers,
    regionalBands,
    marketBands,
    semrushCaps,
    crawlCadence,
    oncrawlCaps,
  };
}

/*************************
 * HELPERS
 *************************/
function inferTier_(revenue, tiers) {
  for (const t of tiers) {
    if (revenue >= t.min && revenue <= t.max) return t;
  }
  return tiers[tiers.length-1] || {name:'Tier D', base:100, ceiling:250};
}
function cadenceFor_(tierName, paying, map) {
  const t = map[tierName] || map['Tier D'] || {nonpaying:'Quarterly', paying:'Quarterly'};
  return paying ? t.paying : t.nonpaying;
}
function readTable_(values, header, rowHandler) {
  let start = -1;
  for (let i=0;i<values.length;i++) {
    if (values[i][0] === header) { start = i+1; break; }
  }
  if (start < 0) return;
  for (let r=start; r<values.length; r++) {
    const row = values[r];
    if (!row[0]) break; // stop at blank separator row
    rowHandler(row);
  }
}
function getLastRow_(sheet, col, startRow) {
  const values = sheet.getRange(startRow, col, sheet.getMaxRows()-startRow+1, 1).getValues();
  for (let i=values.length-1;i>=0;i--) {
    if (String(values[i][0]).trim() !== '') return startRow + i;
  }
  return startRow - 1;
}
function getOrCreateSheet_(ss, name) { let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name); return sh; }
function safeStr_(v){ return (v==null?'':String(v)).trim(); }
function toNumber_(v){ const n = typeof v==='number'?v:parseFloat(String(v).replace(/,/g,'')); return isFinite(n)?n:0; }
function round2_(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }
