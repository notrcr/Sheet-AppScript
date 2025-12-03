// Copy All Sheet w Data/Value on D4 To New Files And Delete A3:F (LINE 21)

function copyAllSheets() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  // Get This Month -1
  const Month = new Date();
  Month.setMonth(Month.getMonth() - 1);
  const MonthName = Month.toLocaleString("en-US", { month: "short"});
  // Make New File N the Id
  const newFile = SpreadsheetApp.create("Hasil Laporan " + MonthName);
  const newSS = SpreadsheetApp.openById(newFile.getId());
  // Loop sheet that D4 is not null then copy Everything to new sheet and delete A3:F 
  sheets.forEach(sh => {
    const d4 = sh.getRange("D4").getValue();  
    
    if (d4 !== "" && d4 !== null) {
      const copied = sh.copyTo(newSS);
      copied.setName(sh.getName()); 
 
      if (sh.getName() === "SheetX") {
        sh.getRange("A3:F").clearContent();
      }
    }
  });
  // Delete The Sheet1 in the New Sheet
  const def = newSS.getSheetByName("Sheet1");
  if (def) newSS.deleteSheet(def);

  SpreadsheetApp.flush();
}
