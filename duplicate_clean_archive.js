function insertDuplicateAndArchive() {
  // constants
  var TABLE_START = 1;
  var TABLE_LENGTH = 15;
  var CHANGABLE_CONTENT_START_ROW= 4;
  var DATE_START_COLUMN = 3;
  var DATE_END_COLUMN = 10;
  var ACTIVE_SHEET = 'Sheet1';
  var ARCHIVE_SHEET = 'Archive';

  // Part 1: Insert and Duplicate
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numRowsToInsert = TABLE_LENGTH;
  var numRowsToDuplicate = TABLE_LENGTH;
  var startRowToDuplicate = TABLE_LENGTH + 1;
  var targetRowForDuplication = TABLE_START;
  var numCols = sheet.getLastColumn();
 
  // Insert blank rows at the specified position
  sheet.insertRows(1, numRowsToInsert);
 
  // Get the range of data to duplicate
  var sourceRange = sheet.getRange(startRowToDuplicate, 1, numRowsToDuplicate, numCols);
 
  // Get the target range for pasting
  var targetRange = sheet.getRange(targetRowForDuplication, 1);
 
  // Copy and paste the data
  sourceRange.copyTo(targetRange, {contentsOnly:true});
 
  // Copy and paste formatting
  var targetRangeEnd = sheet.getRange(targetRowForDuplication + numRowsToDuplicate - 1, numCols);
  sourceRange.copyFormatToRange(sheet, 1, numCols, targetRowForDuplication, targetRowForDuplication + numRowsToDuplicate - 1);
 
  // Clear specific range
  var startRowToClear = CHANGABLE_CONTENT_START_ROW;
  var endRowToClear = TABLE_LENGTH;
  var startColumnToClear = DATE_START_COLUMN; // Column C
  var endColumnToClear = DATE_END_COLUMN; // Column I
  var rangeToClear = sheet.getRange(startRowToClear, startColumnToClear, endRowToClear - startRowToClear + 1, endColumnToClear - startColumnToClear + 1);
  rangeToClear.clearContent();
 
  // Update cells with dates
  var today = new Date();
  var nextSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() + (7 - today.getDay()) % 7);
  var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var dateFormat = 'dd MMM';
  for (var i = 0; i < days.length; i++) {
    sheet.getRange(DATE_START_COLUMN, DATE_START_COLUMN + i).setValue(Utilities.formatDate(nextSunday, Session.getScriptTimeZone(), dateFormat));
    nextSunday.setDate(nextSunday.getDate() + 1);
  }
 
  // Part 3: Copy Rows to Archive and Clear
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(ACTIVE_SHEET); // Assuming "Sheet1" is the sheet to archive from. Adjust if different.
  var archiveSheet = ss.getSheetByName(ARCHIVE_SHEET); // Target sheet name
 
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET);
  }
 
  var startRow = TABLE_LENGTH * 2 + 1; // Adjust the start row if needed
  var numRows = TABLE_LENGTH; // Number of rows to copy
  var numColumns = sourceSheet.getLastColumn();
 
  var rangeToCopy = sourceSheet.getRange(startRow, 1, numRows, numColumns);
  var targetRow = Math.max(archiveSheet.getLastRow() + 1, 1);
 
  rangeToCopy.copyTo(archiveSheet.getRange(targetRow, 1, numRows, numColumns), {contentsOnly: false});
  rangeToCopy.clear(); // Clear content, format, and notes
}
