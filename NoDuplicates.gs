/**
 * Adds a "NoDuplicates" menu to the spreadsheet toolbar when the file opens.
 * "Refresh Now" lets you manually re-run the deduplication at any time.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NoDuplicates')
    .addItem('Refresh Now', 'removeDuplicateQuotes')
    .addSeparator()
    .addItem('Start 1-min auto-refresh', 'setupMinuteTrigger')
    .addItem('Stop 1-min auto-refresh',  'removeMinuteTrigger')
    .addToUi();
}

/**
 * Installs a trigger that runs removeDuplicateQuotes every minute.
 * Use "NoDuplicates > Start 1-min auto-refresh" in the sheet menu,
 * or run it once from the Apps Script editor.
 * Running it again will not create duplicate triggers.
 */
function setupMinuteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'removeDuplicateQuotes' &&
        triggers[i].getTriggerSource() === ScriptApp.TriggerSource.CLOCK &&
        triggers[i].getEventType()    === ScriptApp.EventType.CLOCK) {
      // Check if an existing trigger runs on a 1-minute interval.
      // everyMinutes(1) triggers report a specific type — delete and recreate
      // only if it doesn't already exist at the 1-minute cadence.
      ScriptApp.deleteTrigger(triggers[i]); // replace any old interval with 1 min
    }
  }

  ScriptApp.newTrigger('removeDuplicateQuotes')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert(
    '1-minute auto-refresh started.\n\nUse "Stop 1-min auto-refresh" to cancel it.'
  );
}

/**
 * Removes the 1-minute trigger (and any other time-based trigger for
 * removeDuplicateQuotes). Keeps the 2-hour trigger if you want that instead —
 * run setupTrigger() to reinstate it after stopping the minute trigger.
 */
function removeMinuteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'removeDuplicateQuotes') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  SpreadsheetApp.getUi().alert('Auto-refresh stopped.');
}

/**
 * Reads the top 50 data rows from the "QUOTE-PLEASE" tab, skips rows where
 * the same customer (column F) already appears within the same calendar week
 * (derived from column E), and writes the de-duplicated results to the
 * "NoDuplicates" tab.
 *
 * Called automatically by the active trigger, or manually via
 * NoDuplicates > Refresh Now in the sheet menu.
 */
function removeDuplicateQuotes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Log all sheet names so a name mismatch is immediately visible in the logs.
  var sheetNames = ss.getSheets().map(function(s) { return '"' + s.getName() + '"'; }).join(', ');
  Logger.log('Available sheets: ' + sheetNames);

  var sourceSheet = ss.getSheetByName('SuperJoin QUOTE-PLEASE');
  var destSheet   = ss.getSheetByName('NoDuplicates');

  if (!sourceSheet) {
    Logger.log('ERROR: Sheet "SuperJoin QUOTE-PLEASE" not found. Check the name above.');
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

  // Log header so column positions are visible in the execution log.
  Logger.log('Header row: ' + JSON.stringify(allValues[0]));

  // Log first 3 data rows to confirm what is actually in columns E (4) and F (5).
  for (var d = 0; d < Math.min(3, dataRows.length); d++) {
    Logger.log('Data row ' + (d + 2) + ' col E=' + JSON.stringify(dataRows[d][4]) +
               ' col F=' + JSON.stringify(dataRows[d][5]) +
               ' full=' + JSON.stringify(dataRows[d]));
  }

  var seen       = {};
  var uniqueRows = [];

  for (var i = 0; i < dataRows.length; i++) {
    var row      = dataRows[i];
    var dateValue = row[4];                       // Column E
    var customer  = String(row[5]).trim();        // Column F

    // Skip only rows that are entirely blank across all cells.
    var allBlank = row.every(function(cell) { return cell === '' || cell === null || cell === undefined; });
    if (allBlank) continue;

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

  var date;

  if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    // Column E contains strings like "11/5/2026-21:50:18:354-v31q9".
    // Extract just the leading date portion before the first hyphen.
    var str = String(dateValue);
    var dateMatch = str.match(/^(\d{1,2}\/\d{1,2}\/\d{4})/);
    date = dateMatch ? new Date(dateMatch[1]) : new Date(str);
  }

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
