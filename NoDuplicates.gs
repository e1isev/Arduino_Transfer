/**
 * Reads the top 50 data rows from the "QUOTE-PLEASE" tab, skips rows where
 * the same customer (column F) already appears within the same calendar week
 * (derived from column E), and writes the de-duplicated results to the
 * "NoDuplicates" tab.
 *
 * Called automatically every 2 hours by the trigger installed via
 * setupTrigger(). Can also be run manually from the Apps Script editor.
 */
function removeDuplicateQuotes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Log all sheet names so a name mismatch is immediately visible in the logs.
  var sheetNames = ss.getSheets().map(function(s) { return '"' + s.getName() + '"'; }).join(', ');
  Logger.log('Available sheets: ' + sheetNames);

  var sourceSheet = ss.getSheetByName('QUOTE-PLEASE');
  var destSheet   = ss.getSheetByName('NoDuplicates');

  if (!sourceSheet) {
    Logger.log('ERROR: Sheet "QUOTE-PLEASE" not found. Check the name above.');
    return;
  }

  if (!destSheet) {
    destSheet = ss.insertSheet('NoDuplicates');
    Logger.log('NoDuplicates sheet created.');
  }

  // getDataRange() is safer than getLastColumn() — it never returns 0.
  var allValues = sourceSheet.getDataRange().getValues();
  Logger.log('Rows in QUOTE-PLEASE (including header): ' + allValues.length);

  if (allValues.length < 2) {
    Logger.log('No data rows found — NoDuplicates cleared.');
    destSheet.clearContents();
    return;
  }

  var numCols    = allValues[0].length;
  var headerRow  = [allValues[0]];
  var dataRows   = allValues.slice(1, 51); // top 50 data rows
  Logger.log('Data rows to process: ' + dataRows.length + ' | columns: ' + numCols);

  var seen       = {};
  var uniqueRows = [];

  for (var i = 0; i < dataRows.length; i++) {
    var row      = dataRows[i];
    var dateValue = row[4];                       // Column E
    var customer  = String(row[5]).trim();        // Column F

    if (!dateValue && !customer) continue;        // skip blank rows

    var weekKey   = getWeekStartKey(dateValue);
    var dedupeKey = weekKey + '|' + customer.toLowerCase();

    if (!seen[dedupeKey]) {
      seen[dedupeKey] = true;
      uniqueRows.push(row);
    } else {
      Logger.log('Skipped duplicate — row ' + (i + 2) + ': customer="' + customer + '" week=' + weekKey);
    }
  }

  Logger.log('Unique rows to write: ' + uniqueRows.length);

  destSheet.clearContents();
  destSheet.getRange(1, 1, 1, numCols).setValues(headerRow);

  if (uniqueRows.length > 0) {
    destSheet.getRange(2, 1, uniqueRows.length, numCols).setValues(uniqueRows);
  }

  Logger.log('Done.');
}

/**
 * Installs a time-based trigger that runs removeDuplicateQuotes every 2 hours.
 * Run this function ONCE from the Apps Script editor to activate the schedule.
 * Running it again will not create duplicate triggers.
 */
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'removeDuplicateQuotes') {
      return; // Trigger already exists.
    }
  }

  ScriptApp.newTrigger('removeDuplicateQuotes')
    .timeBased()
    .everyHours(2)
    .create();

  SpreadsheetApp.getUi().alert(
    'Trigger set. removeDuplicateQuotes will run every 2 hours.'
  );
}

/**
 * Removes all time-based triggers for removeDuplicateQuotes.
 * Run this from the Apps Script editor if you want to stop the schedule.
 */
function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'removeDuplicateQuotes') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Returns a string that identifies the Monday of the week containing dateValue.
 * Used as the week portion of the deduplication key.
 *
 * @param {Date|string|number} dateValue  The value from column E.
 * @returns {string}
 */
function getWeekStartKey(dateValue) {
  if (!dateValue) return 'no-date';

  var date = (dateValue instanceof Date) ? dateValue : new Date(dateValue);

  if (isNaN(date.getTime())) return String(dateValue);

  // Shift to the Monday of the same week (Sunday = 0 → previous Monday).
  var dayOfWeek = date.getDay();
  var offsetToMonday = (dayOfWeek === 0) ? -6 : 1 - dayOfWeek;
  var monday = new Date(date);
  monday.setDate(date.getDate() + offsetToMonday);

  return (
    monday.getFullYear() +
    '-' +
    String(monday.getMonth() + 1).padStart(2, '0') +
    '-' +
    String(monday.getDate()).padStart(2, '0')
  );
}
