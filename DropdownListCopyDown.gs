

function dropdownListCopyDown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  // Check C6 Dropdown
  const validation = sheet.getRange("C6").getDataValidation();
  if (!validation) {
    SpreadsheetApp.getUi().alert("C6 has no dropdown.");
    return;
  }
  // Get the Value From the Dropdown Range (The real range that the dropdown reads from.)
  const nameRange = validation.getCriteriaValues()[0];
  const names = nameRange.getValues().flat().filter(n => n);

  if (names.length === 0) {
    SpreadsheetApp.getUi().alert("No names found.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "Copy Data",
    `This will copy the template ${names.length} times. Continue?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  // Check Rows copied and start at 23
  const template = sheet.getRange("A1:G23");
  const rowsPerCopy = template.getNumRows();
  let startRow = 24;

  names.forEach((name, i) => {

    const targetRow = startRow + i * rowsPerCopy;
    const targetRange = sheet.getRange(targetRow, 1, rowsPerCopy, 7);
    // COpy eveerything like normal copypaste down 
    template.copyTo(targetRange, { contentsOnly: false });
 
    sheet.getRange(targetRow + 5, 3).setValue(name);


    Logger.log(`Created block for: ${name} at C${targetRow + 5}`);
  });

  ui.alert(`Completed! ${names.length} blocks created.`);
}
