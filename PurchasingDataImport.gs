function importJFBuyData() {
  const ss = SpreadsheetApp.getActive();
  const SOURCE_ID = '14fpl9TaiJFZrLxem7BZYc9g1kwGmLA1Yf_8lrW3VFUc';
  const SOURCE_TAB_NAME = 'Imported Data - Dont Touch';
  const TARGET_TAB_NAME = 'Import - JF Buy data';

  try {
    const sourceSs = SpreadsheetApp.openById(SOURCE_ID);
    const sourceSh = sourceSs.getSheetByName(SOURCE_TAB_NAME);
    if (!sourceSh) throw new Error(`Source tab "${SOURCE_TAB_NAME}" not found.`);

    const data = sourceSh.getDataRange().getValues();
    if (data.length === 0) throw new Error('Source sheet is empty.');

    const targetSh = getOrCreateSheet_(ss, TARGET_TAB_NAME);
    targetSh.clear();
    targetSh.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Optional: Formatting
    targetSh.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    targetSh.setFrozenRows(1);

    SpreadsheetApp.getUi().alert(`Successfully imported ${data.length} rows into "${TARGET_TAB_NAME}".`);
  } catch (e) {
    console.error('Error importing JF Buy Data:', e);
    SpreadsheetApp.getUi().alert(`Error importing data: ${e.message}`);
  }
}
