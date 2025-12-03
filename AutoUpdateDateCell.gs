  function updateDateInCell() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLokasi");   
    var cell = sheet.getRange("C1");  
    
    var today = new Date(); 
    today.setHours(0, 0, 0, 0); 
    cell.setValue(today);
    cell.setNumberFormat("MM/dd/yyyy");  
}