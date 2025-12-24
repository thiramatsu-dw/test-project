function myFunction() {
    var spreadsheetId = '1nZFiS6wFj3Jk9W7QdKzzMfLkr2PY5hVjre5ZvPoglfc';
    var sheetName = '売上';
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);

    // Get all values in column A
    // We grab up to the max rows in the sheet to check for gaps
    var lastRow = sheet.getMaxRows();
    var values = sheet.getRange(1, 1, lastRow, 1).getValues();

    var targetRow = -1;

    // Find the first empty cell
    for (var i = 0; i < values.length; i++) {
        if (values[i][0] === "") {
            targetRow = i + 1; // 1-based index
            break;
        }
    }

    // If we found an empty row, write to it
    if (targetRow !== -1) {
        sheet.getRange(targetRow, 1).setValue(100);
        Logger.log('Wrote 100 to row ' + targetRow);
    } else {
        // If no empty row found (sheet is full?), insert a new row or just log (unlikely default behavior)
        Logger.log('No empty rows found in existing range.');
        // Optional: sheet.insertRowAfter(lastRow); sheet.getRange(lastRow + 1, 1).setValue(100);
    }
}
