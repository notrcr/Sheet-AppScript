function convertBase64FotoToDriveLink() {
  const sheetName = 'X'; 
  const folderId = 'DriveFolderIdHere';  
  const col = 7; // Column G
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const folder = DriveApp.getFolderById(folderId);

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const timeNow = new Date();

  let count = 0;

  for (let i = 1; i < data.length; i++) {
    // array index so col - 1
    const cellValue = data[i][col -1]; 
    if (typeof cellValue === 'string' && cellValue.startsWith('data:image/')) {
      try {
        // Extract base64 Data with blob decode then create file in Drive free access and view link
        const base64Data = cellValue.split(,)[1] || cellValue;
          const blob = Utilities.newblob(Utilities.base64Decode(base64Data), 'image/jpeg',  `Foto_${i}_${timeNow.getTime()}.jpg`);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        // g\et the link n set to cell
        const link = file.getUrl(); 
        sheet.getRange(i + 1, col).setValue(link);
        count++;
      }
      catch (e) { 
        console.error(`Error processing row ${i + 1}: ${e}`);
      }
    }
  }
   Logger.log(`✅ Converted ${count} Base64 images to Drive links.`);

}