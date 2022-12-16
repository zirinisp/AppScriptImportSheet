function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Data Updates')
    .addItem('Update All Data Sources', 'importDataSources')
    .addItem('Update Current Sheet Data Sources', 'importCurrentSheetDataSource')
          .addSeparator()
          .addSubMenu(ui.createMenu('Update Single Data Source')
              .addItem('Missing Items Data', 'updateMissingItems')
              .addItem('Ratings Data', 'updateRatingsData')
              .addItem('Statement Import Data', 'updateStatementImportData')
              .addItem('Staff Data', 'updateStaffData'))
    .addToUi();
}

function updateMissingItems() {
  importDataSources(["Missing Items Data"]);
}

function updateRatingsData() {
  importDataSources(["Ratings Data"]);
}
function updateStatementImportData() {
  importDataSources(["Statement Import Data"]);
}
function updateStaffData() {
  importDataSources(["Staff Data"]);
}


var OnHoldSpreadsheetNameCell = "Orders";
var OnHoldRangeSpreadsheetCell = "A3";
var UpdateSpreadsheetNames =  [
    "Missing Items Data",
    "Ratings Data",
    "Statement Import Data",
    "Staff Data"
  ];

function importDataSources(spreadsheetNames = UpdateSpreadsheetNames) {

  var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var onHoldSheet = destinationSpreadsheet.getSheetByName(OnHoldSpreadsheetNameCell);

  destinationSpreadsheet.toast("Pausing Updates");

  onHoldSheet.getRange(OnHoldRangeSpreadsheetCell).setValue("On Hold");
  SpreadsheetApp.flush();
  destinationSpreadsheet.toast("Updates Paused");
  for (var i = 0; i < spreadsheetNames.length; i++) {
    var destinationSheetName = spreadsheetNames[i];
    importRangeFor(destinationSpreadsheet, destinationSheetName, false);
    SpreadsheetApp.flush();
  }

  destinationSpreadsheet.toast("Resuming Updates");
  onHoldSheet.getRange(OnHoldRangeSpreadsheetCell).setValue("Running");
  SpreadsheetApp.flush();
  destinationSpreadsheet.toast("Updates Resumed");
}

function importCurrentSheetDataSource() {

  var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var destinationSheetName = destinationSpreadsheet.getActiveSheet().getName();
  importDataSources([destinationSheetName]);

}

var SourceRangeSpreadsheetCell = "B1"
var SourceRangeSheetNameCell = "C1"
var SourceRangeColumnCell = "D1";
var DestinationFirstRow = 3;

function importRangeFor(destinationSpreadsheet, destinationSheetName, flush = true) {
  destinationSpreadsheet.toast("Importing " + destinationSheetName);
  var destinationSS = destinationSpreadsheet.getSheetByName(destinationSheetName);

  var sourceSpreadsheetId = destinationSS.getRange(SourceRangeSpreadsheetCell).getValue();
  var sourceSheetName = destinationSS.getRange(SourceRangeSheetNameCell).getValue();
  var sourceRangeColumn = destinationSS.getRange(SourceRangeColumnCell).getValue();
  var numOfColumns = destinationSS.getMaxColumns();
  var sourceColumns = destinationSS.getRange(2, 1, 1, numOfColumns).getValues();
  var sourceStartRow = sourceRangeColumn;

  if (sourceSpreadsheetId.startsWith("https://")) {
    sourceSpreadsheetId = sourceSpreadsheetId.substring(
      sourceSpreadsheetId.indexOf("/d/") + 3,
      sourceSpreadsheetId.lastIndexOf("/")
    );
  }

  const sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);

  for (var i = 0; i < sourceColumns[0].length; i++) {
    var sourceColumn = sourceColumns[0][i];
    var logMessage = "Importing " + destinationSheetName + " " + i + "/" + sourceColumns[0].length + "\ni: " + i + " column: " + sourceColumn;
    Logger.log(logMessage);

    if (sourceColumn != "") {
      destinationSpreadsheet.toast(logMessage);
      var sourceRange = sourceSheetName + "!" + sourceColumn + sourceStartRow + ":" + sourceColumn;
      var destinationColumnLetter = columnToLetter(i + 1);
      var destinationRangeStart = destinationSheetName + "!" + destinationColumnLetter + DestinationFirstRow;

      importRange(sourceSS, sourceRange, destinationSS, destinationRangeStart, flush);
    }
  }

}

/**
* Imports range data from one Google Sheet to another.
* @param {string} sourceID - The id of the source Google Sheet.
* @param {string} sourceRange - The Sheet tab and range to copy.
* @param {string} destinationID - The id of the destination Google Sheet.
* @param {string} destinationRangeStart - The destintation location start cell as a sheet name and cell.
*/
function importRange(sourceSS, sourceRange, destinationSS, destinationRangeStart, flush = true) {

  // Gather the source range values
  const sourceRng = sourceSS.getRange(sourceRange)
  const sourceVals = sourceRng.getValues();

  // Get the destiation sheet and cell location.
  const destStartRange = destinationSS.getRange(destinationRangeStart);
  const destSheet = destStartRange.getSheet();

  // Clear previous entries.
  //destSheet.clear();

  // Get the full data range to paste from start range.
  const destRange = destSheet.getRange(
    destStartRange.getRow(),
    destStartRange.getColumn(),
    sourceVals.length,
    sourceVals[0].length
  );

  var numOfRowsToClear = destSheet.getMaxRows() - destStartRange.getRow();

  // Get the range that need to be cleared before writing.
  const destClearRange = destSheet.getRange(
    destStartRange.getRow(),
    destStartRange.getColumn(),
    numOfRowsToClear,
    sourceVals[0].length
  );

  destClearRange.clearContent();

  // Paste in the values.
  destRange.setValues(sourceVals);

  if (flush) {
    SpreadsheetApp.flush();
  }
};

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}