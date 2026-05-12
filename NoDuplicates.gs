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
  var sourceSheet = ss.getSheetByName('QUOTE-PLEASE');
  var destSheet = ss.getSheetByName('NoDuplicates');

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('Sheet "QUOTE-PLEASE" not found.');
    return;
  }

  if (!destSheet) {
    destSheet = ss.insertSheet('NoDuplicates');
  }

  var lastCol = sourceSheet.getLastColumn();

  // Row 1 is assumed to be the header row.
  var headerValues = sourceSheet.getRange(1, 1, 1, lastCol).getValues();

  // Grab the top 50 data rows (rows 2–51).
  var dataRows = sourceSheet.getRange(2, 1, 50, lastCol).getValues();

  var seen = {};
  var uniqueRows = [];

  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];

    // Column E = index 4 (date / week), Column F = index 5 (customer).
    var dateValue = row[4];
    var customer = String(row[5]).trim();

    // Skip entirely blank rows.
    if (!dateValue && !customer) continue;

    var weekKey = getWeekStartKey(dateValue);
    var dedupeKey = weekKey + '|' + customer.toLowerCase();

    if (!seen[dedupeKey]) {
      seen[dedupeKey] = true;
      uniqueRows.push(row);
    }
  }

  // Overwrite the destination tab with the filtered data.
  destSheet.clearContents();
  destSheet.getRange(1, 1, 1, lastCol).setValues(headerValues);

  if (uniqueRows.length > 0) {
    destSheet.getRange(2, 1, uniqueRows.length, lastCol).setValues(uniqueRows);
  }

  // Only show the alert when run manually (triggers have no UI context).
  try {
    SpreadsheetApp.getUi().alert(
      uniqueRows.length + ' unique row(s) written to the NoDuplicates tab.'
    );
  } catch (e) {
    // Running via trigger — no UI available, silently continue.
  }
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
