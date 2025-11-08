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
    const accountKey = r.account.toLowerCase();
    const techFee = techByAccount.get(accountKey) || 0;
    const paying  = techFee > 0;

    // If inactive, force to "Inactive" with zero base/ceiling
    const tierObj = inactive ? { name:'Inactive', base:0, ceiling:0 } : inferTier_(revenue, cfg.tiers);

    const bandName = cfg.marketBands[r.market] || 'Low';
    const regionMult = cfg.regionalBands[bandName] != null ? cfg.regionalBands[bandName] : 1.00;

    const siteSize = cfg.accountSiteSizes[accountKey] || 'Default';
    const siteMultiplier = cfg.siteSizeMultipliers && cfg.siteSizeMultipliers.hasOwnProperty(siteSize)
      ? cfg.siteSizeMultipliers[siteSize]
      : (cfg.siteSizeMultipliers && cfg.siteSizeMultipliers.Default != null ? cfg.siteSizeMultipliers.Default : 1);
    const ownCrawler = !!cfg.accountCrawlerOptOut[accountKey];
    const bonusRaw = ownCrawler ? (cfg.ownCrawlerAccuMult || 1) : 1;
    const bonusSafe = (isFinite(bonusRaw) && bonusRaw > 0) ? bonusRaw : 1;
    return {
      account: r.account,
      market:  r.market,
      revenue,
      techFee,
      paying,
      inactive,                 // <— new flag
      tier: tierObj.name,
      tierBase: tierObj.base,
      accuBonus: bonusSafe,
      regionBandName: bandName,
      regionMult,
      siteSize,
      siteMultiplier: isFinite(siteMultiplier) ? siteMultiplier : 1,
      ownCrawler
    };
  });
  // --- AccuRanker allocation (deterministic by tier/size/bonuses)
  const accuCapacity = Math.max(0, cfg.accuCapacity || 0);
  const techFeeBonus = isFinite(cfg.techFeeBonus) ? cfg.techFeeBonus : 0;
  const nonPayerMult = (isFinite(cfg.nonPayerMult) && cfg.nonPayerMult > 0) ? cfg.nonPayerMult : 1;

  rows.forEach(r => {
    if (r.inactive) {
      r.accuBase = 0;
      r.accuContributor = 0;
      r.accuTotal = 0;
      r.techBonusApplied = 0;
      return;
    }

    const tierBase = Math.max(0, r.tierBase || 0);
    const regionalFactor = r.regionMult || 1;
    const siteFactor = r.siteMultiplier != null ? r.siteMultiplier : 1;
    const crawlerFactor = r.accuBonus || 1;
    const payerMultiplier = r.paying ? 1 : nonPayerMult;
    const payerBonus = r.paying ? techFeeBonus : 0;

    let allocation = tierBase * regionalFactor * siteFactor;
    allocation = allocation * payerMultiplier;
    allocation += payerBonus;
    allocation = allocation * crawlerFactor;

    r.accuBase = Math.max(0, Math.round(allocation));
    r.accuContributor = 0;
    r.accuTotal = r.accuBase;
    r.techBonusApplied = payerBonus;
  });

  const totalAccuAllocated = rows.reduce((s, r) => s + (r.accuTotal || 0), 0);
  const accuRemaining = accuCapacity - totalAccuAllocated;
  const accuUtilPercent = accuCapacity > 0 ? round2_((totalAccuAllocated / accuCapacity) * 100) : 0;


  // --- Semrush caps (inactive = 0)
  rows.forEach(r => {
    if (r.inactive) { r.semrushCap = 0; return; }
    const caps = cfg.semrushCaps[r.tier] || cfg.semrushCaps['Default'] || { nonpaying: 50, paying: 100 };
    // use explicit fallbacks for each side and fix typo
    r.semrushCap = r.paying ? (caps.paying ?? 100) : (caps.nonpaying ?? 50);
  });


  // --- OnCrawl cadence (inactive = None)
  rows.forEach(r => {
    if (r.inactive) {
      r.oncrawlCadence = 'None';
    } else if (r.ownCrawler && cfg.ownCrawlerSkipOncrawl) {
      r.oncrawlCadence = 'Client crawler';
    } else {
      r.oncrawlCadence = cadenceFor_(r.tier, r.paying, cfg.crawlCadence);
    }
  });

  // --- OnCrawl starter caps (inactive = 0)
  rows.forEach(r => {
    if (r.inactive || (r.ownCrawler && cfg.ownCrawlerSkipOncrawl)) {
      r.oncrawlBase = 0;
    } else {
      const caps = cfg.oncrawlCaps[r.tier] || cfg.oncrawlCaps['Tier D'] || { nonpaying: 2500, paying: 4000 };
      const baseCap = r.paying ? (caps.paying ?? 4000) : (caps.nonpaying ?? 2500);
      const multiplier = r.siteMultiplier != null ? r.siteMultiplier : 1;
      r.oncrawlBase = Math.round(baseCap * multiplier);
    }
  });

  // Output
  const outName = `Tech Fair-Usage — ${year}`;
  const outSh = getOrCreateSheet_(ss, outName);
  const header = [
    'Account','Market','Year',
    'Tier','Site Size','Own Crawler?','Pays Tech Fee?','Revenue','Tech Fee',
    'Regional Band','Tech Fee Bonus (Keywords)',
    'AccuRanker Base','AccuRanker Contributor','AccuRanker Total',
    'OnCrawl Base','OnCrawl Contributor','OnCrawl Total',
    'Semrush Keyword Cap','OnCrawl Cadence'
  ];
  const out = [header];
  rows.forEach(r => out.push([
    r.account, r.market, year,
    r.tier, r.siteSize, r.ownCrawler ? 'Yes':'No', r.paying ? 'Yes':'No', r.revenue, r.techFee,
    r.regionBandName, r.techBonusApplied || 0,
    r.accuBase, r.accuContributor, r.accuTotal,
    // OnCrawl uses starter caps for now. Contributor logic can be added later.
    r.oncrawlBase, 0, r.oncrawlBase,
    r.semrushCap, r.oncrawlCadence
  ]));

  outSh.clearContents();
  outSh.getRange(1,1,out.length,header.length).setValues(out);

  const summaryData = [
    ['Accu Capacity', accuCapacity],
    ['Accu Allocated', totalAccuAllocated],
    ['Accu Remaining', accuRemaining],
    ['Accu Utilization %', `${accuUtilPercent}%`]
  ];
  const summaryStartCol = header.length + 3;
  outSh.getRange(1, summaryStartCol, summaryData.length, 2).setValues(summaryData);
  outSh.getRange(1, summaryStartCol, summaryData.length, 2).setFontWeight('bold');

  // Formatting
  outSh.getRange(1,1,1,header.length).setFontWeight('bold');
  if (out.length > 1) {
    outSh.getRange(2,8,out.length-1,2).setNumberFormat('#,##0');      // revenue, tech fee
    outSh.getRange(2,11,out.length-1,1).setNumberFormat('#,##0');     // tech fee bonus
    outSh.getRange(2,12,out.length-1,3).setNumberFormat('#,##0');     // AccuRanker numbers
    outSh.getRange(2,15,out.length-1,3).setNumberFormat('#,##0');     // OnCrawl numbers
    outSh.getRange(2,18,out.length-1,1).setNumberFormat('#,##0');     // Semrush cap
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
function EnsureSetupTab_(options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActive();
  const shName = 'Setup';
  let sh = ss.getSheetByName(shName);
  if (!sh) sh = ss.insertSheet(shName);

  if (sh.getLastRow() < 5) {
    renderSetupSheet_(sh, defaultSetupRenderState_());
  }

  let values = sh.getDataRange().getDisplayValues();
  if (!values || values.length === 0) {
    renderSetupSheet_(sh, defaultSetupRenderState_());
    values = sh.getDataRange().getDisplayValues();
  }

  const firstCol = values.map(r => r[0] || '');
  const val = (key, def) => {
    const row = values.find(r => r[0] === key);
    return row ? row[1] : def;
  };

  // Allow a typo safety net from previous version
  const fallbackCap = parseInt(val('ACCUNOTREAL', '100000'), 10);
  const capacity = parseInt(val('ACCURANKER_CAPACITY', fallbackCap || 100000), 10);
  const techFeeBonus = parseFloat(val('TECH_FEE_BONUS', 250));
  const nonPayerMult = parseFloat(val('NON_PAYER_MULT', 0.85));
  const ownCrawlerAccuMult = parseFloat(val('OWN_CRAWLER_ACCU_MULT', 1.25));
  const ownCrawlerSkipOncrawl = parseYesNo_(val('OWN_CRAWLER_SKIP_ONCRAWL', 'Yes'));

  // Parse “Revenue Tiers” block with “Base / Ceiling” column
  const tiers = [];
  readTable_(values, ['Client Tier Matrix','Revenue Tiers'], row => {
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
  readTable_(values, ['Regional Band Multipliers','Regional Bands'], row => {
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

  // Site size multipliers
  const siteSizeMultipliers = {};
  readTable_(values, ['Website Size Multipliers','Site Size Multipliers'], row => {
    if (row[0]) {
      const key = String(row[0]).trim();
      const multiplier = parseFloat(row[1]);
      if (key) siteSizeMultipliers[key] = isFinite(multiplier) ? multiplier : 1;
    }
  });

  // Account level site size assignments
  const accountSiteSizes = {};
  readTable_(values, 'Account→Site Size', row => {
    const account = row[0] ? String(row[0]).trim() : '';
    const size = row[1] ? String(row[1]).trim() : '';
    if (account && size) accountSiteSizes[account.toLowerCase()] = size;
  });

  // Account level crawler opt-outs
  const accountCrawlerOptOut = {};
  readTable_(values, 'Account→Crawler Status', row => {
    const account = row[0] ? String(row[0]).trim() : '';
    const setting = row[1] ? String(row[1]).trim() : '';
    if (!account) return;
    if (parseYesNo_(setting)) {
      accountCrawlerOptOut[account.toLowerCase()] = true;
    }
  });

  const capturedTables = {
    tiers: captureTable_(values, ['Client Tier Matrix','Revenue Tiers']),
    semrush: captureTable_(values, 'Semrush Caps'),
    onCrawl: captureTable_(values, 'OnCrawl Starter Caps'),
    cadence: captureTable_(values, 'Crawl Cadence Rules'),
    regionalBands: captureTable_(values, ['Regional Band Multipliers','Regional Bands']),
    marketBands: captureTable_(values, 'Market→Regional Band'),
    siteSizeMultipliers: captureTable_(values, ['Website Size Multipliers','Site Size Multipliers']),
    accountSizes: captureTable_(values, 'Account→Site Size'),
    crawlerStatus: captureTable_(values, 'Account→Crawler Status'),
    allocationGuidance: captureTable_(values, 'Allocation Guidance')
  };

  const needsRefresh = ['Client Tier Matrix','Regional Band Multipliers','Website Size Multipliers','Account→Crawler Status','Allocation Guidance','OWN_CRAWLER_ACCU_MULT','TECH_FEE_BONUS','NON_PAYER_MULT']
    .some(label => firstCol.indexOf(label) === -1);

  if (needsRefresh && !opts._skipRender) {
    renderSetupSheet_(sh, {
      keyValues: {
        ACCURANKER_CAPACITY: capacity || 100000,
        TECH_FEE_BONUS: isFinite(techFeeBonus) ? techFeeBonus : 250,
        NON_PAYER_MULT: (isFinite(nonPayerMult) && nonPayerMult > 0) ? nonPayerMult : 0.85,
        OWN_CRAWLER_ACCU_MULT: isFinite(ownCrawlerAccuMult) ? ownCrawlerAccuMult : 1.25,
        OWN_CRAWLER_SKIP_ONCRAWL: ownCrawlerSkipOncrawl ? 'Yes' : 'No'
      },
      tableRows: capturedTables
    });
    return EnsureSetupTab_({ _skipRender: true });
  }

  return {
    accuCapacity: capacity,
    tiers,
    regionalBands,
    marketBands,
    semrushCaps,
    crawlCadence,
    oncrawlCaps,
    siteSizeMultipliers,
    accountSiteSizes,
    accountCrawlerOptOut,
    ownCrawlerAccuMult: isFinite(ownCrawlerAccuMult) && ownCrawlerAccuMult > 0 ? ownCrawlerAccuMult : 1,
    techFeeBonus: isFinite(techFeeBonus) ? techFeeBonus : 0,
    nonPayerMult: (isFinite(nonPayerMult) && nonPayerMult > 0) ? nonPayerMult : 1,
    ownCrawlerSkipOncrawl,
  };
}

function renderSetupSheet_(sheet, state) {
  const keyValues = (state && state.keyValues) || {};
  const tableRows = (state && state.tableRows) || {};
  const rows = [];
  const pushBlank = () => rows.push(['','','','']);

  rows.push(['Key','Value','Notes','']);
  rows.push(padRow4_(['ACCURANKER_CAPACITY', keyValues.ACCURANKER_CAPACITY ?? 100000, 'AccuRanker capacity (≈100k tracking slots).','']));
  rows.push(padRow4_(['TECH_FEE_BONUS', keyValues.TECH_FEE_BONUS ?? 250, 'Flat keyword bonus awarded to tech-fee payers.','']));
  rows.push(padRow4_(['NON_PAYER_MULT', keyValues.NON_PAYER_MULT ?? 0.85, 'Multiplier applied to non-paying accounts (e.g., 0.85 = 15% reduction).','']));
  rows.push(padRow4_(['OWN_CRAWLER_ACCU_MULT', keyValues.OWN_CRAWLER_ACCU_MULT ?? 1.25, 'Multiplier applied to AccuRanker base + ceiling when a client has their own crawler.','']));
  rows.push(padRow4_(['OWN_CRAWLER_SKIP_ONCRAWL', keyValues.OWN_CRAWLER_SKIP_ONCRAWL ?? 'Yes', '"Yes" disables OnCrawl cadence/URLs when Own Crawler? = true.','']));

  pushBlank();
  rows.push(['Client Tier Matrix','Revenue Min','Revenue Max','Accu Base / Ceiling']);
  (tableRows.tiers && tableRows.tiers.length ? tableRows.tiers : defaultTierRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Semrush Caps','Non-paying','Paying','(per client)']);
  (tableRows.semrush && tableRows.semrush.length ? tableRows.semrush : defaultSemrushRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['OnCrawl Starter Caps','Non-paying','Paying','(monthly starter defaults)']);
  (tableRows.onCrawl && tableRows.onCrawl.length ? tableRows.onCrawl : defaultOncrawlRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Crawl Cadence Rules','Non-paying','Paying','(OnCrawl)']);
  (tableRows.cadence && tableRows.cadence.length ? tableRows.cadence : defaultCadenceRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Regional Band Multipliers','Multiplier','','Notes']);
  (tableRows.regionalBands && tableRows.regionalBands.length ? tableRows.regionalBands : defaultRegionalBandRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Market→Regional Band','Band','','Notes']);
  (tableRows.marketBands && tableRows.marketBands.length ? tableRows.marketBands : defaultMarketBandRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Website Size Multipliers','Multiplier','Page Count Guidance','Notes']);
  (tableRows.siteSizeMultipliers && tableRows.siteSizeMultipliers.length ? tableRows.siteSizeMultipliers : defaultSiteSizeRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Account→Site Size','Site Size','Pages (optional)','Notes']);
  (tableRows.accountSizes && tableRows.accountSizes.length ? tableRows.accountSizes : defaultAccountSizeRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Account→Crawler Status','Own Crawler?','Notes','']);
  (tableRows.crawlerStatus && tableRows.crawlerStatus.length ? tableRows.crawlerStatus : defaultCrawlerRows_())
    .forEach(row => rows.push(padRow4_(row)));

  pushBlank();
  rows.push(['Allocation Guidance','AccuRanker Keywords','Semrush Keywords','OnCrawl URLs / Cadence']);
  (tableRows.allocationGuidance && tableRows.allocationGuidance.length ? tableRows.allocationGuidance : defaultAllocationRows_())
    .forEach(row => rows.push(padRow4_(row)));

  sheet.clear();
  sheet.getRange(1,1,rows.length,4).setValues(rows);
  sheet.setFrozenRows(1);
  for (let c=1;c<=4;c++) sheet.autoResizeColumn(c);
}

function defaultSetupRenderState_() {
  return {
    keyValues: {
      ACCURANKER_CAPACITY: 100000,
      TECH_FEE_BONUS: 250,
      NON_PAYER_MULT: 0.85,
      OWN_CRAWLER_ACCU_MULT: 1.25,
      OWN_CRAWLER_SKIP_ONCRAWL: 'Yes'
    },
    tableRows: {
      tiers: defaultTierRows_(),
      semrush: defaultSemrushRows_(),
      onCrawl: defaultOncrawlRows_(),
      cadence: defaultCadenceRows_(),
      regionalBands: defaultRegionalBandRows_(),
      marketBands: defaultMarketBandRows_(),
      siteSizeMultipliers: defaultSiteSizeRows_(),
      accountSizes: defaultAccountSizeRows_(),
      crawlerStatus: defaultCrawlerRows_(),
      allocationGuidance: defaultAllocationRows_()
    }
  };
}

function defaultTierRows_() {
  return [
    ['Tier A',500000,'','800 / 2000'],
    ['Tier B',200000,499999,'500 / 1200'],
    ['Tier C',50000,199999,'250 / 600'],
    ['Tier D',0,49999,'100 / 250']
  ];
}

function defaultSemrushRows_() {
  return [
    ['Tier A',200,400,''],
    ['Tier B',150,300,''],
    ['Tier C',100,200,''],
    ['Tier D',50,100,'']
  ];
}

function defaultOncrawlRows_() {
  return [
    ['Tier A',25000,40000,''],
    ['Tier B',10000,18000,''],
    ['Tier C',5000,8000,''],
    ['Tier D',2500,4000,'']
  ];
}

function defaultCadenceRows_() {
  return [
    ['Tier A','Monthly','Weekly/Fortnightly',''],
    ['Tier B','Bi-monthly or Quarterly','Monthly',''],
    ['Tier C','Quarterly','Quarterly',''],
    ['Tier D','One-off / Quarterly by request','One-off / Quarterly by request','']
  ];
}

function defaultRegionalBandRows_() {
  return [
    ['Top',1.20,'','High focus markets'],
    ['Mid',1.10,'','Developed markets'],
    ['Low',1.00,'','Baseline']
  ];
}

function defaultMarketBandRows_() {
  return [
    ['UK','Top','',''],
    ['ZA','Top','',''],
    ['US','Mid','',''],
    ['FR','Low','','']
  ];
}

function defaultSiteSizeRows_() {
  return [
    ['Default',1.0,'< 5k pages or standard demand','Fallback when no size is specified'],
    ['Small',0.8,'< 5k indexed pages','Lower crawl demand'],
    ['Medium',1.0,'5k – 30k indexed pages','Baseline allocation'],
    ['Large',1.3,'30k+ indexed pages or ecommerce','Boosted allocation']
  ];
}

function defaultAccountSizeRows_() {
  return [
    ['Example Account A','Large','','Replace with your account + size'],
    ['Example Account B','Medium','',''],
    ['Example Account C','Small','','']
  ];
}

function defaultCrawlerRows_() {
  return [
    ['Example Account D','Yes','Skips OnCrawl and gains Accu bonus','']
  ];
}

function defaultAllocationRows_() {
  return [
    ['How to use','Accu = Tier base × regional multiplier × site-size multiplier (own crawler bonus applies before pool).','Semrush caps follow tier table; apply regional multipliers if required.','OnCrawl starter caps × site multiplier; set to zero when Own Crawler? = Yes.'],
    ['Example (Tier B • Mid • Large)','≈500 × 1.10 × 1.30 ≈ 715 base keywords before contributor pool.','150 / 300 × 1.10 ≈ 165 / 330 keywords.','10k / 18k × 1.30 ≈ 13k / 23k URLs; cadence from rules.']
  ];
}

function padRow4_(row) {
  return [row[0] ?? '', row[1] ?? '', row[2] ?? '', row[3] ?? ''];
}

function captureTable_(values, header) {
  const headers = Array.isArray(header) ? header : [header];
  for (let i=0; i<values.length; i++) {
    if (headers.indexOf(values[i][0]) !== -1) {
      const rows = [];
      for (let r=i+1; r<values.length; r++) {
        const row = values[r];
        if (!row[0]) break;
        rows.push(padRow4_(row));
      }
      return rows;
    }
  }
  return [];
}

function parseYesNo_(value) {
  if (value === undefined || value === null) return false;
  return /^(yes|y|true|1)$/i.test(String(value).trim());
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
  const headers = Array.isArray(header) ? header : [header];
  let start = -1;
  for (let i=0;i<values.length;i++) {
    if (headers.indexOf(values[i][0]) !== -1) { start = i+1; break; }
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
