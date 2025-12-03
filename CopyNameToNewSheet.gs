// 1. Source Sheet
const SOURCE_SHEET_NAME = 'Karyawan';

// 2. List Name to be pasted
const BASE_SHEET_NAMES  = [
    "A","B","C","D","E",];


function copyDataAndFormulas() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    ui.alert(`Error: Source sheet "${SOURCE_SHEET_NAME}" not found.`);
    return;
  }
  let copiedCount = 0;
  // Loop through each base name 
  for (const baseName of BASE_SHEET_NAMES) {

    const sheetName = `${baseName} `;

    //If sheet exists delete it and recreate from template
    let oldSheet = ss.getSheetByName(sheetName);
    if (oldSheet) ss.deleteSheet(oldSheet);

    // Duplicate template again (fresh)
    let destinationSheet = sourceSheet.copyTo(ss);
    destinationSheet.setName(sheetName);

    // Put name in B1
    destinationSheet.getRange("B1").setValue(baseName);

    copiedCount++;
  }

  ui.alert(`${copiedCount} sheets Done.`);
}

