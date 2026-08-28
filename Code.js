/**
 * Expense Tracker — Google Apps Script backend
 * Bind this script to your "Expense-Tracker" spreadsheet
 * (Extensions > Apps Script from inside the sheet).
 */

var EXPENSE_SHEET = 'Expenses'; // type: expense
var INCOME_SHEET = 'credit';    // type: income

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found. Check the sheet name.');
  }
  return sheet;
}

function formatDateForClient_(date) {
  if (Object.prototype.toString.call(date) === '[object Date]' && !isNaN(date)) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return date || '';
}

/**
 * Reads both sheets and returns a combined, sorted list of transactions.
 */
function getTransactions() {
  var configs = [
    { name: EXPENSE_SHEET, type: 'expense' },
    { name: INCOME_SHEET, type: 'income' }
  ];
  var result = [];

  configs.forEach(function (cfg) {
    var sheet = getSheet_(cfg.name);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    values.forEach(function (row, idx) {
      var rowNum = idx + 2;
      var date = row[0];
      var description = row[1];
      var notes = row[2];
      var amount = row[3];

      // Skip fully blank rows
      if (!description && !amount) return;

      result.push({
        id: cfg.name + '_' + rowNum,
        date: formatDateForClient_(date),
        description: description || '',
        notes: notes || '',
        amount: Number(amount) || 0,
        type: cfg.type
      });
    });
  });

  result.sort(function (a, b) {
    return new Date(b.date) - new Date(a.date);
  });

  return result;
}

/**
 * Appends a transaction to the correct sheet based on tx.type,
 * then returns the refreshed full list.
 */
function addTransaction(tx) {
  if (!tx || !tx.description || !tx.amount || !tx.date) {
    throw new Error('Missing required fields.');
  }
  var sheetName = tx.type === 'income' ? INCOME_SHEET : EXPENSE_SHEET;
  var sheet = getSheet_(sheetName);
  var dateValue = new Date(tx.date + 'T00:00:00');

  sheet.appendRow([dateValue, tx.description, tx.notes || '', Number(tx.amount)]);
  return getTransactions();
}

/**
 * Deletes a transaction by its composite id ("SheetName_RowNumber"),
 * then returns the refreshed full list.
 */
function deleteTransaction(id) {
  if (!id) throw new Error('Missing id.');
  var lastUnderscore = id.lastIndexOf('_');
  var sheetName = id.substring(0, lastUnderscore);
  var rowNum = parseInt(id.substring(lastUnderscore + 1), 10);

  if (!sheetName || isNaN(rowNum)) {
    throw new Error('Invalid transaction id.');
  }

  var sheet = getSheet_(sheetName);
  sheet.deleteRow(rowNum);
  return getTransactions();
}
