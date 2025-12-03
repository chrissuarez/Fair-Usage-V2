/**
 * Email Import Utilities
 * Handles importing CSV data from Gmail attachments based on labels.
 */

/**
 * Generic helper to import the *latest* CSV attachment from a thread with the given label.
 * @param {string} labelName - Gmail label to search for (e.g. "label:dashboard-reports...").
 * @param {string} targetSheetName - Name of the sheet to overwrite.
 * @param {Object} options - Optional settings:
 *   - mapFunction: (row, headers) => transformedRow
 *   - requiredHeaders: Array of strings to validate/force headers.
 *   - startRow: Row index to start reading from (default 1).
 */
function importCsvFromGmailLabel_(labelName, targetSheetName, options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, targetSheetName);
  
  try {
    const threads = GmailApp.search(labelName, 0, 1); // Get latest 1
    if (threads.length === 0) {
      console.warn(`No threads found for label: ${labelName}`);
      return `No emails found for ${labelName}`;
    }
    
    const messages = threads[0].getMessages();
    const lastMsg = messages[messages.length - 1]; // Get latest message in thread
    const attachments = lastMsg.getAttachments();
    
    const csvAttachment = attachments.find(a => a.getContentType() === 'text/csv' || a.getName().endsWith('.csv'));
    if (!csvAttachment) {
      console.warn(`No CSV attachment found in latest email for ${labelName}`);
      return `No CSV found for ${labelName}`;
    }
    
    const csvData = Utilities.parseCsv(csvAttachment.getDataAsString('ISO-8859-1'));
    if (csvData.length < 1) return `Empty CSV for ${labelName}`;
    
    // Clear and Write
    sh.clear();
    
    // If we have a map function or specific header requirements, process here.
    // For now, we'll just dump the raw data as requested, but we can add transformation if needed.
    
    if (csvData.length > 0) {
      sh.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    }
    
    return `Imported ${csvData.length} rows to ${targetSheetName}`;
    
  } catch (e) {
    console.error(`Error importing ${labelName}: ${e.message}`);
    // Don't throw, just return error string so pipeline continues? 
    // Or throw to stop pipeline? User probably wants to know.
    throw new Error(`Failed to import ${labelName}: ${e.message}`);
  }
}

/**
 * 1. Opportunities and Accounts
 * Label: label:dashboard-reports-opportunities-and-accounts
 * Target: Estimate_to_Opportunity_Map
 */
function importOpportunitiesAndAccounts() {
  return importCsvFromGmailLabel_(
    'label:dashboard-reports-opportunities-and-accounts', 
    'Estimate_to_Opportunity_Map'
  );
}

/**
 * 2. Projects and Opportunity Lookup
 * Label: label:dashboard-reports-projects-and-opportunity-lookup
 * Target: Projects_RAW_Data_Import
 */
function importProjectsAndOppLookup() {
  return importCsvFromGmailLabel_(
    'label:dashboard-reports-projects-and-opportunity-lookup', 
    'Projects_RAW_Data_Import'
  );
}

/**
 * 3. Opps and CRs
 * Label: label:dashboard-reports-opps-and-crs
 * Target: Opps_and_CRs_RAW_Import
 */
function importOppsAndCRs() {
  return importCsvFromGmailLabel_(
    'label:dashboard-reports-opps-and-crs', 
    'Opps_and_CRs_RAW_Import'
  );
}
