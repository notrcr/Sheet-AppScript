// Lock calculated numeric values in the range B4:AF58 of sheet "X"

function lockCalculatedValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("X");  
  // Check Sheet
  if (!sheet) {
    console.log("X tidak ditemukan.");
    return;
  }

  // Fixed range B4:AF58
  const startRow = 4;
  const startCol = 2;   // Column B
  const numRows = 55;   // 58 - 4 + 1
  const numCols = 32;    
  // Get the Range , Formulas , Values
  const dataRange = sheet.getRange(startRow, startCol, numRows, numCols);
  const formulas = dataRange.getFormulas();
  const values = dataRange.getValues();

  let processedCount = 0;
  // if still have data loop i and j
  for (let i = 0; i < formulas.length; i++) {
    for (let j = 0; j < formulas[i].length; j++) {

      const row = startRow + i;
      const col = startCol + j;
      const hasFormula = formulas[i][j];
      const currentValue = values[i][j];
      // If it is formula with type of number set the value .
      if (hasFormula && typeof currentValue === 'number' && !isNaN(currentValue)) {
        sheet.getRange(row, col).setValue(currentValue);  
        processedCount++;
      }
    }
  }

  console.log(`Sukses! ${processedCount} nilai berhasil dikunci dalam B4:AF58.`);
}
