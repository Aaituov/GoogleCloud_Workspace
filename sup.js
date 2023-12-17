function removeDuplicatesAndCopyToNewSheet() {
  var sourceSheetName = "Base";
  var targetSheetName = "test";

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);

  // Assuming 'email' is the unique identifier
  var uniqueIdentifierColumnIndex = getColumnIndexByName(sourceSheet, "Your student ID");

  if (uniqueIdentifierColumnIndex === -1) {
    Logger.log("Column not found");
    return;
  }

  // Get data from the sheet
  var data = getDataFromSheet(sourceSheet);

  // Identify unique entries based on 'email'
  var uniqueEntries = removeDuplicatesByColumn(data, uniqueIdentifierColumnIndex);

  // Create or clear the target sheet
  var targetSheet = createOrClearSheet(spreadsheet, targetSheetName);

  // Write unique entries to the target sheet
  targetSheet.getRange(1, 1, uniqueEntries.length, uniqueEntries[0].length).setValues(uniqueEntries);
}

function getColumnIndexByName(sheet, columnName) {
  var headerRow = getDataFromSheet(sheet)[0];

  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === columnName) {
      return i;
    }
  }

  // Return -1 if the column is not found
  return -1;
}

function removeDuplicatesByColumn(data, columnIndex) {
  var uniqueEntries = [];
  var uniqueValues = new Set();

  for (var i = 0; i < data.length; i++) {
    var value = data[i][columnIndex];

    // Check if the value is unique
    if (!uniqueValues.has(value)) {
      uniqueValues.add(value);
      uniqueEntries.push(data[i]);
    }
  }

  return uniqueEntries;
}

// The rest of your functions remain the same
