function resetFormResponsesDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("X");

  // Create a date object n set yesterday
  const today = new Date();
  today.setDate(today.getDate() - 1);
  const date = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
 
  const backupName = "X" + date;
  // Create a new sheet for the backup
  const backupSheet = ss.insertSheet(backupName);

  // Get all values only from the form sheet
  const dataRange = formSheet.getDataRange();
  const data = dataRange.getValues();

  // Paste the data values into the backup sheet
  backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  const formats = dataRange.getNumberFormats();
  backupSheet.getRange(1, 1, formats.length, formats[0].length).setNumberFormats(formats);

  // Copy visual formatting (colors, borders, alignment, etc.)
  dataRange.copyFormatToRange(
    backupSheet,
    1,
    dataRange.getLastColumn(),
    1,
    dataRange.getLastRow()
  );

  // Match widths n height with the original sheet
  for (let i = 1; i <= formSheet.getLastColumn(); i++) {
    backupSheet.setColumnWidth(i, formSheet.getColumnWidth(i));
  }
  for (let i = 1; i <= formSheet.getLastRow(); i++) {
    backupSheet.setRowHeight(i, formSheet.getRowHeight(i));
  }
  // Delete all rows except the header row
  const lastRow = formSheet.getLastRow();
  if (lastRow > 1) {
    const numRows = lastRow - 1;      
    formSheet.deleteRows(2, numRows);  
  }
 
  Logger.log("Backup created and form reset successfully.");
}
//  + Update date in C1
function updateDateInCell() {
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLokasi");
  var cell = sheet.getRange("C1");
  cell.setValue(new Date());
}
