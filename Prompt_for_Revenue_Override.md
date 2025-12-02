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
2. **Pass 2 (Generate Rows):**  
   * In the main loop where masterRows are created, check if a Scaling Factor exists for the current row's Account+Opportunity.  
   * If it exists, multiply the calculated rawLineAmount (Hours \* Rate) by the Scaling Factor *before* it is assigned to totalAmount.  
   * Add a note to the Debug\_Info array: "Scaled by \[X\]%" so I can verify it.

Specific Code Implementation:  
Here is the logic I want you to integrate. Please replace the existing rebuildMasterForSource\_ function with this updated version that includes the pre-calculation logic:  
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

  // \--- NEW LOGIC: PASS 1 (Calculate Scaling Factors) \---  
  const scalingFactors \= {};   
    
  if (cfg \=== SOURCE\_CONFIG.Estimate) {  
    const fm \= cfg.fieldMap;  
    const oppTotals \= {}; // { "Account::OppName": 700000 }  
      
    // 1\. Sum up the raw amounts per Opportunity Group  
    rawRows.forEach(raw \=\> {  
      const acc \= safeStr\_(raw\[headerMap\[fm.account\]\]);  
      const opp \= safeStr\_(raw\[headerMap\[fm.opportunity\]\]);  
      const key \= \`${acc}::${opp}\`;   
        
      const hours \= toNumber\_(raw\[headerMap\[fm.hours\]\]);  
      const rate \= toNumber\_(raw\[headerMap\[fm.billRate\]\]);  
        
      if (\!oppTotals\[key\]) oppTotals\[key\] \= 0;  
      oppTotals\[key\] \+= (hours \* rate);  
    });

    // 2\. Check for "Target\_Revenue" overrides  
    Object.keys(oppTotals).forEach(key \=\> {  
      const parts \= key.split('::');  
      const groupUid \= generateOpportunityUid\_(parts\[0\], parts\[1\]);   
      const rawTotal \= oppTotals\[key\];

      // Check for Target\_Revenue override on the Group UID  
      if (overrides\[groupUid\] && overrides\[groupUid\]\['target\_revenue'\]) {  
         const target \= Number(overrides\[groupUid\]\['target\_revenue'\]);  
         if (\!isNaN(target) && rawTotal \> 0\) {  
           scalingFactors\[key\] \= target / rawTotal;   
         }  
      }  
    });  
  }  
  // \--- END NEW LOGIC \---

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

    if (cfg \=== SOURCE\_CONFIG.Estimate) {  
      const fm \= cfg.fieldMap;  
    
      account \= safeStr\_(get(fm.account));  
      oppName \= safeStr\_(get(fm.opportunity));  
      const estId \= safeStr\_(get(fm.estimateId));  
      uid \= estId || generateOpportunityUid\_(account, oppName);  
        
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
        
      // Calculate RAW Amount  
      const hours \= toNumber\_(get(fm.hours));  
      const billRate \= toNumber\_(get(fm.billRate));  
      let rawLineAmount \= hours \* billRate;

      // \--- NEW LOGIC: Apply Scaling Factor \---  
      const groupKey \= \`${account}::${oppName}\`;  
      if (scalingFactors\[groupKey\]) {  
        const factor \= scalingFactors\[groupKey\];  
        rawLineAmount \= rawLineAmount \* factor;  
        debugInfo.push(\`Scaled by ${(factor\*100).toFixed(1)}%\`);  
      }  
      // \---------------------------------------

      totalAmount \= rawLineAmount;  
      capability \= categorizeRevenue(get(fm.role));  
      techFeePaying \= false;

    } else {  
      // Tech Fee Logic (Unchanged)  
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

    // Existing Override logic (Specific Row Overrides take precedence over scaling)  
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