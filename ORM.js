
function getDuplicateLocationsInColumnA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const map = {};

  sheets.forEach(sheet => {
    const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();  // Only Column A
    const name = sheet.getName();

    for (let r = 0; r < values.length; r++) {
      const val = values[r][c];  // c value is an integer and should be fixed
      if (val !== "") {
        const key = val.toString().trim();
        const loc = `${name}!R${r + 1}C{c + 1}`;  // Format of the array can be different 
        // accordingly
        
        if (!map[key]) map[key] = [];
        map[key].push(loc);
      }
    }
  });

  const resultSheet = ss.getSheetByName('Duplicate Report') || ss.insertSheet('Duplicate Report');
  resultSheet.clearContents().appendRow(['Value', 'Occurrences', 'Locations']);

  for (const key in map) {
    if (map[key].length > 1) {
      resultSheet.appendRow([key, map[key].length, map[key].join(', ')]);
    }
  }
}
