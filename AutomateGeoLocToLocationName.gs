function convertGeoToLocation() {
  const sheetName = 'X';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const addressCell = sheet.getRange(row, 9); // Column I = 9
    const address = addressCell.getValue();
    
    // Skip  if blank
    if (address && address !== '') {
      continue;
    }
    // Get coordinate from Column F
    const coord = sheet.getRange(row, 6).getValue(); // Column F = 6
    if (coord && typeof coord === 'string' && coord.includes(',')) {
      try {
        // split lat lng and reverse geocode    
        const [lat, lng] = coord.split(',').map(Number);
        const response = Maps.newGeocoder().reverseGeocode(lat, lng);
        const results = response.results;
        const addressResult = results && results.length > 0 ? results[0].formatted_address : 'Not found';
        addressCell.setValue(addressResult);
      } catch (e) {
        addressCell.setValue('Error');
      }
    }
  }
}