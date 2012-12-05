/*global SpreadsheetApp: false, Browser: false */
function listSheets() {
    'use strict';
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = activeSpreadsheet.getSheets(),
        sheetCount = sheets.length,
        i;
    for (i = 0; i < sheetCount; i += 1) {
        Browser.msgBox(sheets[i].getName());
    }
}
