I need you to act as a Senior Google Apps Script Developer. I need to update the file TechRevenueOps.gs to handle "Proportional Revenue Overrides" for my Opportunity data.

The Problem:  
Currently, my code pulls in revenue data from Salesforce (Estimates) which consists of multiple rows per Opportunity (different roles, hours, etc.). Sometimes the total Salesforce revenue is wrong. I want to override the Total Revenue for an Opportunity (e.g., set it to $200k) and have the script automatically proportionally scale all the individual resource rows for that opportunity so they sum up to exactly $200k, preserving their relative weights.  
The Task:  
Please modify the rebuildMasterForSource\_ function in TechRevenueOps.gs. You need to implement a "Two-Pass" approach:

1. **Pass 1 (Calculate Scaling Factors):**  
   * Iterate through all rawRows first.  
   * Sum the raw revenue (Hours \* Bill Rate) for every Opportunity Group (grouped by Account \+ Opportunity Name).  
   * Check the overrides object. If there is an override for the Group UID with the field Target\_Revenue (case-insensitive), calculate a **Scaling Factor** (Override Target / Raw Total).  
   * Store this factor in a map: scalingFactors\[groupKey\] \= factor.  
2. **Pass 2 (Generate Rows & Map New Columns):**  
   * Extract the new columns: Resource Role, Pricing Region (as Region), Hours, and Bill Rate.  
   * Apply the Scaling Factor to the revenue calculation if it exists.  
   * Pass these new fields into the masterRows object so they can be printed.  
3. **Update Column Headers:**  
   * Update the writeMasterSheet\_ function to use the specific column order: Opportunity\_UID, Account, Opportunity Name, Capability, Resource Role, Resource Region, Hours, Bill Rate, Start Date, End Date, Total\_USD, Tech\_Fee\_Paying?, followed by the months.

Specific Code Implementation:  
Please replace the rebuildMasterForSource\_ and writeMasterSheet\_ functions with this code:  
function rebuildMasterForSource\_(cfg, globalCfg, dateLookup, projectLookup) {  
  const ss \= SpreadsheetApp.getActive();  
  const rawSh \= ss.getSheetByName(cfg.rawSheet);  
  const ovSh \= ss.getSheetByName(cfg.overrideSheet);  
  const masterSh \= ss.getSheetByName(cfg.masterSheet);  
  if (\!rawSh || \!ovSh || \!masterSh) throw new Error(\`Missing shadow table for ${cfg.rawSheet}\`);  
    
  const headers \= rawSh.getRange(1, 1, 1, rawSh.getLastColumn()).getDisplayValues()\[0\];  
  const rawRows \= rawSh.getLastRow() \> 1 ?  
    rawSh.getRange(2, 1, rawSh.getLastRow() \- 1, headers.length).getValues() : \[\];  
  const overrides \= readOverrides\_(ovSh);

  const monthsSet \= new Set();  
  const masterRows \= \[\];  
  const headerMap \= headers.reduce((acc, h, idx) \=\> { acc\[h\] \= idx; return acc; }, {});

  // \--- PASS 1: Calculate Scaling Factors \---  
  const scalingFactors \= {};   
    
  if (cfg \=== SOURCE\_CONFIG.Estimate) {  
    const fm \= cfg.fieldMap;  
    const oppTotals \= {}; // { "Account::OppName": 700000 }  
      
    // Sum up the raw amounts per Opportunity Group  
    rawRows.forEach(raw \=\> {  
      const acc \= safeStr\_(raw\[headerMap\[fm.account\]\]);  
      const opp \= safeStr\_(raw\[headerMap\[fm.opportunity\]\]);  
      const key \= \`${acc}::${opp}\`;   
        
      const hours \= toNumber\_(raw\[headerMap\[fm.hours\]\]);  
      const rate \= toNumber\_(raw\[headerMap\[fm.billRate\]\]);  
        
      if (\!oppTotals\[key\]) oppTotals\[key\] \= 0;  
      oppTotals\[key\] \+= (hours \* rate);  
    });

    // Check for "Target\_Revenue" overrides  
    Object.keys(oppTotals).forEach(key \=\> {  
      const parts \= key.split('::');  
      const groupUid \= generateOpportunityUid\_(parts\[0\], parts\[1\]);   
      const rawTotal \= oppTotals\[key\];

      if (overrides\[groupUid\] && overrides\[groupUid\]\['target\_revenue'\]) {  
         const target \= Number(overrides\[groupUid\]\['target\_revenue'\]);  
         if (\!isNaN(target) && rawTotal \> 0\) {  
           scalingFactors\[key\] \= target / rawTotal;   
         }  
      }  
    });  
  }

  // \--- PASS 2: Generate Rows \---  
  rawRows.forEach(raw \=\> {  
    const get \= name \=\> raw\[headerMap\[name\]\];  
    let uid \= '';  
    let account \= '';  
    let oppName \= '';  
    let startDate \= null;  
    let endDate \= null;  
    let currency \= 'USD';  
    let totalAmount \= 0;  
    let capability \= '';  
    let techFeePaying \= false;  
    let debugInfo \= \[\];  
      
    // New Fields  
    let role \= '';  
    let region \= '';  
    let hours \= 0;  
    let billRate \= 0;

    if (cfg \=== SOURCE\_CONFIG.Estimate) {  
      const fm \= cfg.fieldMap;  
    
      account \= safeStr\_(get(fm.account));  
      oppName \= safeStr\_(get(fm.opportunity));  
      const estId \= safeStr\_(get(fm.estimateId));  
      uid \= estId || generateOpportunityUid\_(account, oppName);  
        
      // Extract new columns  
      role \= safeStr\_(get(fm.role));  
      region \= safeStr\_(get(fm.region));  
      hours \= toNumber\_(get(fm.hours));  
      billRate \= toNumber\_(get(fm.billRate));

      // Project Lookup & Dates Logic  
      if (projectLookup && projectLookup\[oppName\]) {  
        startDate \= projectLookup\[oppName\].start;  
        endDate \= projectLookup\[oppName\].end;  
        debugInfo.push(\`Dates from Project: ${projectLookup\[oppName\].name}\`);  
      }  
      if (\!startDate) {  
        startDate \= parseDate\_(get(fm.startDate));  
        endDate \= parseDate\_(get(fm.endDate));  
      }  
        
      currency \= safeStr\_(get(fm.currency)) || 'USD';  
        
      let rawLineAmount \= hours \* billRate;

      // Apply Scaling Factor  
      const groupKey \= \`${account}::${oppName}\`;  
      if (scalingFactors\[groupKey\]) {  
        const factor \= scalingFactors\[groupKey\];  
        rawLineAmount \= rawLineAmount \* factor;  
        debugInfo.push(\`Scaled by ${(factor\*100).toFixed(1)}%\`);  
      }

      totalAmount \= rawLineAmount;  
      capability \= categorizeRevenue(get(fm.role));  
      techFeePaying \= false;

    } else {  
      // Tech Fee Logic  
      const fm \= cfg.fieldMap;  
      uid \= safeStr\_(get(fm.opportunityUid)) || generateOpportunityUid\_(safeStr\_(get(fm.account)), safeStr\_(get(fm.opportunity)));  
      account \= safeStr\_(get(fm.account));  
      oppName \= safeStr\_(get(fm.opportunity));  
        
      if (projectLookup && projectLookup\[oppName\]) {  
        startDate \= projectLookup\[oppName\].start;  
        endDate \= projectLookup\[oppName\].end;  
        debugInfo.push(\`Dates from Project: ${projectLookup\[oppName\].name}\`);  
      }  
      if (\!startDate && dateLookup && dateLookup\[oppName\]) {  
        startDate \= dateLookup\[oppName\].start;  
        endDate \= dateLookup\[oppName\].end;  
        debugInfo.push('Dates from Estimate Opp');  
      }  
      if (\!startDate) {  
        if (fm.startDate) startDate \= parseDate\_(get(fm.startDate));  
        if (\!startDate && fm.closeDate) startDate \= parseDate\_(get(fm.closeDate));  
        if (startDate) {  
           if (fm.endDate) endDate \= parseDate\_(get(fm.endDate));  
           if (\!endDate) {  
             const d \= new Date(startDate);  
             d.setFullYear(d.getFullYear() \+ 1);  
             d.setDate(d.getDate() \- 1);  
             endDate \= d;  
             debugInfo.push('Dates Defaulted (1yr)');  
           }  
        }  
      }  
      currency \= safeStr\_(get(fm.currency)) || 'USD';  
      totalAmount \= parseCurrency\_(get(fm.productAmount));  
      capability \= safeStr\_(get(fm.productName)) || 'Tech Fee';  
      techFeePaying \= true;  
      const stage \= safeStr\_(get(fm.stage)).trim().toLowerCase();  
      if (\!techFeePaying) debugInfo.push(\`Not Paying: Stage="${stage}", Amt=${totalAmount}\`);  
    }

    const applied \= applyOverridesToRow\_(  
      {  
        Opportunity\_UID: uid,  
        Account: account,  
        Opportunity\_Name: oppName,  
        Start\_Date: startDate,  
        End\_Date: endDate,  
        Currency: currency,  
        Total\_Amount: totalAmount,  
        Capability: capability,  
        Tech\_Fee\_Paying: techFeePaying  
      },  
      overrides\[uid\]  
    );  
    if (applied) debugInfo.push('Override Applied');

    const rate \= globalCfg.currencyMap\[currency\] || 1;  
    const totalUsd \= totalAmount \* rate;  
    const monthlyMode \= (globalCfg.params.partialMode || 'SIMPLE').toUpperCase();  
    const monthly \= calculateMonthlyRevenue(totalUsd, startDate, endDate, monthlyMode);  
    Object.keys(monthly).forEach(k \=\> monthsSet.add(k));

    masterRows.push({  
      Opportunity\_UID: uid,  
      Account: account,  
      Opportunity\_Name: oppName,  
      Capability: capability || (cfg.amountField \=== 'Revenue' ? 'Other/Shared' : 'Tech Fee'),  
      Resource\_Role: role,     // NEW  
      Resource\_Region: region, // NEW  
      Hours: hours,            // NEW  
      Bill\_Rate: billRate,     // NEW  
      Start\_Date: startDate,  
      End\_Date: endDate,  
      Total\_USD: totalUsd,  
      Monthly: monthly,  
      Currency: currency,  
      Tech\_Fee\_Paying: techFeePaying,  
      Override\_Applied: applied,  
      Debug\_Info: debugInfo.join('; ')  
    });  
  });

  const months \= Array.from(monthsSet).sort();  
  writeMasterSheet\_(masterSh, masterRows, months, cfg.amountField);

  return { rows: masterRows, months };  
}

function writeMasterSheet\_(sheet, rows, months, amountField) {  
  // Updated Header Layout  
  const header \= \[  
    'Opportunity\_UID',   
    'Account',   
    'Opportunity Name',   
    'Capability',   
    'Resource Role',    // NEW  
    'Resource Region',  // NEW  
    'Hours',            // NEW  
    'Bill Rate',        // NEW  
    'Start Date',   
    'End Date',   
    'Total\_USD',   
    'Tech\_Fee\_Paying?'  
  \].concat(months);

  sheet.clearContents();  
  if (rows.length \=== 0\) {  
    sheet.getRange(1, 1, 1, header.length).setValues(\[header\]).setFontWeight('bold');  
    return;  
  }  
    
  const values \= rows.map(r \=\> {  
    const base \= \[  
      r.Opportunity\_UID,  
      r.Account,  
      r.Opportunity\_Name,  
      r.Capability,  
      r.Resource\_Role,    // NEW  
      r.Resource\_Region,  // NEW  
      r.Hours,            // NEW  
      r.Bill\_Rate,        // NEW  
      r.Start\_Date,  
      r.End\_Date,  
      r.Total\_USD,  
      r.Tech\_Fee\_Paying ? 'Yes' : 'No'  
    \];  
    months.forEach(m \=\> base.push(r.Monthly\[m\] || 0));  
    return base;  
  });

  sheet.getRange(1, 1, values.length \+ 1, header.length).setValues(\[header\].concat(values));  
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');  
  sheet.setFrozenRows(1);  
  // Optional: Auto-resize might be slow for many columns, consider commenting out if slow  
  // for (let c \= 1; c \<= header.length; c++) sheet.autoResizeColumn(c);  
}  