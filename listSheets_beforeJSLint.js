function listSheets () {
   var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
   var sheets = activeSpreadsheet.getSheets();
   for ( var i = 0, sheetCount = sheets.length; i < sheetCount; i++ ) {
       Browser.msgBox(sheets[i].getName());
   }
}
