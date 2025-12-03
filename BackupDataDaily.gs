 function backupDataDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("X");
  
  // Create a date object n set yesterday
  const today = new Date();
  today.setDate(today.getDate() - 1);
  const date = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");

  const backupName = "X" + date;
  if (ss.getSheetByName(backupName)) {
    Logger.log("Backup already exists for today.");
    return;
  }

  const backupSheet = ss.insertSheet(backupName);

  // Copy All Values
  const dataRange = formSheet.getDataRange();
  const data = dataRange.getValues();
  backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
 
  const formats = dataRange.getNumberFormats();
  backupSheet.getRange(1, 1, formats.length, formats[0].length).setNumberFormats(formats);
 
  dataRange.copyFormatToRange(
    backupSheet,
    1,
    dataRange.getLastColumn(),
    1,
    dataRange.getLastRow()
  );
 // For Every Column n Width set same
  for (let i = 1; i <= formSheet.getLastColumn(); i++) {
    backupSheet.setColumnWidth(i, formSheet.getColumnWidth(i));
  } 
  for (let i = 1; i <= formSheet.getLastRow(); i++) {
    backupSheet.setRowHeight(i, formSheet.getRowHeight(i)); 
  }

  Logger.log("Backup created successfully w full formatting.");
}
 

