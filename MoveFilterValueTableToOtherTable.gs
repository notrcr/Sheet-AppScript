// Move Values from Table H:M to A:F in the Same Sheet "X" Without Deleting Existing Data

function MoveFilterValue() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("X");
  if (!sheet) return;
  // Find Last Row
  const startRow = 3;
  const numCols = 6; // H:M → A:F
  // Find Last Row (Has to have data)
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  // Source (The Table to Copy) H:M .  Getting Range From Start Row , Col 8 , Total Row, Total Col) . Get the Value
  const sourceRange = sheet.getRange(startRow, 8, lastRow - startRow + 1, numCols);
  let values = sourceRange.getValues();

  // Remove empty rows (Each row Join Into Single String and if not blank saved it)
  values = values.filter(r => r.join("") !== "");
  if (values.length === 0) return;

  // Find Range from Col A3 until last row and check values
  const colAValues = sheet.getRange(3, 1, sheet.getLastRow()).getValues();
  // Paste Row Start at 3
  let pasteRow = 3;
  // Find First Empty Row in A
  while (pasteRow - 3 < colAValues.length && colAValues[pasteRow - 3].join("") !== "") {
    pasteRow++;
  }

  // Paste values
  sheet.getRange(pasteRow, 1, values.length, numCols).setValues(values);

  SpreadsheetApp.getUi().alert("Saved " + values.length + " rows from H:M → A:F");
}
