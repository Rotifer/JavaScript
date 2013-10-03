/*
All the code examples from "Google Spreadsheet Programming"
Author: Michael Maguire
*/

// Chapter 2

function sayHelloBrowser() {
  // Declare a string literal variable.
  var greeting = 'Hello world!';
  // Display a message dialog with the greeting 
//(visible from the containing spreadsheet).
  Browser.msgBox(greeting);
}
function helloDocument() {
  var greeting = 'Hello world!';
  // Create DocumentApp instance.
  var doc = 
    DocumentApp.create('test_DocumentApp');
  // Write the greeting to a Google document.
  doc.setText(greeting);
  // Close the newly created document
  doc.saveAndClose();  
}
function helloLogger() {
  var greeting = 'Hello world!';
  //Write the greeting to a logging window.
  // This is visible from the script editor
  //   window menu "View->Logs...".
  Logger.log(greeting);  
}
function helloSpreadsheet() {
  var greeting = 'Hello world!',
      sheet = SpreadsheetApp.getActiveSheet();
  // Post the greeting variable value to cell A1
  // of the active sheet in the containing 
  //  spreadsheet.
  sheet.getRange('A1').setValue(greeting);
  // Using the LanguageApp write the 
  //  greeting to cell:
  // A2 in Spanish, 
  //  cell A3 in German, 
  //  and cell A4 in French.
  sheet.getRange('A2')
        .setValue(LanguageApp.translate(
                  greeting, 'en', 'es'));
  sheet.getRange('A3')
        .setValue(LanguageApp.translate(
                  greeting, 'en', 'de'));
  sheet.getRange('A4')
         .setValue(LanguageApp.translate(
                   greeting, 'en', 'fr'));
}


// Chapter 3

// Cannot be called as a UDF.
function setRangeFontBold (rangeAddress) {
  var sheet = 
    SpreadsheetApp.getActiveSheet();
  sheet.getRange(rangeAddress)
    .setFontWeight('bold');
}
// Call "setRangeFontBold()" from editor.
function call_setCellFontBold () {
  var rangeAddress = Browser.inputBox(
                   
                    'Set Range Font Bold', 
                    'Provide a range address',
                     Browser.Buttons.OK_CANCEL);
  if (rangeAddress) {
    setRangeFontBold(rangeAddress);
  }
}
// Given the standard deviation and the mean,
//  return the relative standard deviation.
function RSD (stdev, mean) {
  if (!(typeof stdev === 'number' && 
        typeof mean === 'number')) {
    throw {'name': 'TypeError',
           'message': 
          'Function "RSD()" requires ' +
          'two numeric arguments'};
  }
    return (100 * (stdev/mean)).toFixed(2)*1; 
}
// Given a temperature value in Celsius
//  return the temperature in Fahrenheit.
function celsiusToFahrenheit (celsius) {
if (typeof celsius !== 'number') {
    throw {
      'name': 'TypeError',
      'message': 'Function requires ' +
          'a single number argument'};
  }
  return ((celsius * 9) / 5) + 32;
}
// Given a temperature in Fahrenheit,
// return the temperature in Celsius.
function fahrenheitToCelsius(fahrenheit) {
  if (typeof fahrenheit !== 'number') {
    throw {
      'name': 'TypeError',
      'message': 'Function requires ' +
           ' a single number argument'};
  }
  return ( fahrenheit - 32 ) * 5/9;
}
// Given the radius, return the
// area.
// Throw an error if the radius is
// negative.
function areaOfCircle (radius)  {
  if (typeof radius !== 'number'){
    throw {
      'name': 'TypeError',
      'message': 'Function requires ' + 
        'a single numeric argument'};
  }    
  if (radius < 0) {
    throw {
      'name': 'ValueError',
      'message': 'Radius myst ' +
        ' be non-negative'};    
  }
  return Math.PI * (radius * radius);
}
function test_intervalInDays() {
  var date1 = new Date(),
      date2 = new Date(1972, 7, 17);
  Logger.log(intervalInDays(date1, date2));
}
// Write String methods to the logger.
function printStringMethods() {
  var strMethods = 
    Object.getOwnPropertyNames(
      String.prototype);
  Logger.log('String has ' +
              strMethods.length +
             ' properties.');
  Logger.log(strMethods.sort().join('\n'));
}
// Reverse the alphabet.
function test_reverseString () {
  var str = 'abcdefghijklmnopqrstuvwxyz';
  Logger.log(reverseString(str));
}
// Return a string with the characters

// of the input string reversed.
function reverseString (str) {
  var strReversed = '',
      lastCharIndex = str.length - 1,
      i;
  if (typeof str !== 'string') {
    throw {
      'name': 'TypeError',
      'message': 'Function requires a ' +
        
      ' single string argument.'};
  }
  for (i = lastCharIndex; i >= 0; i -= 1) {
    strReversed += str[i];
  }
  return strReversed;
}
// Return a integer between
// 1 and 6 inclusive.
function throwDie () {
  return 1 + Math.floor(Math.random() * 6);
}
// Concatenate cell values from
// an input range.
// Single quotes around concatenated 
// elements are optional.
function concatRng(inputFromRng, concatStr,
                    addSingleQuotes) {
  var cellValues;
  if (addSingleQuotes) {
    cellValues = 
      inputFromRng.map(
        function (element) {
          return "'" + element + "'";
        });
    return cellValues.join(concatStr);
 }
   return inputFromRng.join(concatStr);
}
// Print stockInfo object property
// names to the logger.
function printFinanceAppKeys() {
  stockSymbol = 'GOOG';
  Logger.log(Object.keys(
            FinanceApp.getStockInfo(
                   stockSymbol))
                   .sort()
                   .join('\n'));
}
// Given a stock symbol, return the
//  stock price (NYSE).
function getStockPrice(stockSymbol) {
  return FinanceApp
   .getStockInfo(stockSymbol)['price'];
}
// Given a stock symbol, return the
//  full stock name. 
function getStockName(stockSymbol) {
  return FinanceApp
   .getStockInfo(stockSymbol)['name'];
}

// Chapter 4

// Function to demonstrate the Spreadsheet 
//  object hierarchy.
// All the variables are gathered in a 
//  JavaScript array.
// At each iteration of the for loop the 
//  "toString()" method 
//  is called for each variable and its
//   output is printed to the log.
function showGoogleSpreadsheetHierarchy() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(),
     sh = ss.getActiveSheet(),
     rng = ss.getRange('A1:C10'),
     innerRng = rng.getCell(3, 3),
     innerRngAddress = innerRng.getA1Notation(),
     column = innerRngAddress.slice(0,1),
     googleObjects = [ss, sh, rng, innerRng, 
                       innerRngAddress, column],
     i;
 for (i = 0; i < googleObjects.length; i += 1) {
   Logger.log(googleObjects[i].toString());
 }
}
// Print the column letter of the third row and 
//   third column of the range "A1:C10"
//  of the active sheet in the active 
//   spreadsheet.
// This is for demonstration purposes only!
function getColumnLetter () {
  Logger.log(
    SpreadsheetApp.getActiveSpreadsheet()
      .getActiveSheet().getRange('A1:C10')
      .getCell(3, 3).getA1Notation()
        .slice(0,1));
}
// Extract an array of all the property names 
//  defined for Spreadsheet and write them to
//  column A of the active sheet in the active
//   spreadsheet.
function testSpreadsheet () {
  var ss = 
     SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet(),
      i,
      spreadsheetProperties = [],
      outputRngStart = sh.getRange('A2');
  sh.getRange('A1')
    .setValue('spreadsheet_properties');
  sh.getRange('A1')
    .setFontWeight('bold');
  spreadsheetProperties = 
    Object.keys(ss).sort();
  for (i = 0; 
       i < spreadsheetProperties.length;
       i += 1) {
    outputRngStart.offset(i, 0)
       .setValue(spreadsheetProperties[i]);
  }
}
//  Extract, an array of properties from a
//   Sheet object.
// Sort the array alphabetically using the
//  Array sort() method.
// Use the Array join() method to a create
//   a string of all the Sheet properties
//   separated by a new line.
function printSheetProperties () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet();
  Logger.log(Object.keys(sh)
             .sort().join('\n'));
} 
// Call function listSheets() passing it the 
//  Spreadsheet object for the active 
//    spreadsheet.
// The try - catch construct handles the 
//  error thrown by listSheets() if the given
// argument is absent or something
//    other than a Spreadsheet object.
function test_listSheets () {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet();
  try {
    listSheets(ss);
  } catch (error) {
    Logger.log(error.message);
  }
}
// Given a Spreadsheet object, 
//  print the names of its sheets
//   to the logger.
// Throw an error if the argument
//   is missing or if it is not
//  of type Spreadsheet.
function listSheets (spreadsheet) {
  var sheets,
      i;
  if (spreadsheet.toString()
      !== 'Spreadsheet') {
    throw {
      'name': 'TypeError',
      'message': 'Function "listSheets()" '
                 'requires ' +
                 'a single argument of ' +
                 'type "Spreadsheet".'};
  }
  sheets = spreadsheet.getSheets();
  for (i = 0; 
       i < sheets.length; i += 1) {
    Logger.log(sheets[i].getName());
  }
}
// Create a Spreadsheet object and call 
//  "sheetExists()" for an array of sheet 
//  names to see if they exist in 
//  the given spreadsheet.
// Print the output to the log.
function test_sheetExists () {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sheetNames = ['Sheet1', 
                    'sheet1', 
                    'Sheet2',
                    'SheetX'],
      i;
  for (i = 0; 
       i < sheetNames.length; 
       i +=1) {
    Logger.log('Sheet Name ' + 
                sheetNames[i] + 
               ' exists: ' + 
                sheetExists(ss, 
                      sheetNames[i]));
  }
}
// Given a Spreadsheet object and a sheet name, 
//  check for two arguments of the correct type.
// Return "true" if the given sheet name exists
//  in the given Spreadsheet, 
//  else return "false".
function sheetExists(spreadsheet, sheetName) {
  var sheet;
  if (spreadsheet.toString() !==
      'Spreadsheet') {
    throw {
      'name': 'TypeError',
      'message': 'Function "sheetExists()" ' +
                 'first argument for ' + 
                 '"spreadsheet" is ' +
                 'not type "Spreadsheet".'};
  }
  if (typeof sheetName !== 'string') {
    throw {
      'name': 'TypeError',
      'message': 'Function "sheetExists()" ' +
                 'second argument ' + 
                 'for "sheetName" ' +
                 'is not type string.'};    
  }
  if (spreadsheet.getSheetByName(sheetName)) {
    return true;
  } else {
    return false;
  }
}
// Copy the first sheet of the active
//  spreadsheet to a newly created 
//  spreadsheet.
function copySheetToSpreadsheet () {
  var ssSource = 
     SpreadsheetApp.getActiveSpreadsheet(),
      ssTarget = 
      SpreadsheetApp.create(
        'CopySheetTest'),
      sourceSpreadsheetName =
        ssSource.getName(),
      targetSpreadsheetName = 
        ssTarget.getName();
  Logger.log(
     'Copying the first sheet from ' + 
            sourceSpreadsheetName + 
            ' to ' + targetSpreadsheetName);
  // [0] extracts the first Sheet object 
  //   from the array created by
  //   method call "getSheets()"
  ssSource.getSheets()[0].copyTo(ssTarget);
}
// Create a Sheet object and pass it 
// as an argument to getSheetSummary().
// Print the return value to the log.
function test_getSheetSummary () {
  var sheet = SpreadsheetApp
             .getActiveSpreadsheet()
             .getSheets()[0];
  Logger.log(getSheetSummary(sheet));
}
// Given a Sheet object as an argument, 
//  use Sheet methods to extract 
//  information about it.
// Collect this information into an object
// literal and return the object literal.
function getSheetSummary (sheet) {
  var sheetReport = {};
  if (sheet.toString() !== 'Sheet') {
    throw {
      'name': 'TypeError',
      'message': 'Function "getSheetReport()" ' +
                'requires a single ' + 
               'argument of type "Sheet".'};
  }
  sheetReport['Sheet Name'] = 
     sheet.getName();
  sheetReport['Used Row Count'] =
     sheet.getLastRow();
  sheetReport['Used Column count'] = 
    sheet.getLastColumn();
  sheetReport['Used Range Address'] = 
      'A1:' + 
   sheet.getRange(sheet.getLastRow(), 
   sheet.getLastColumn()).getA1Notation();
  return sheetReport;
}

// Chapter 5

// Select a number of cells in a spreadsheet and 
//  then execute the following function.
// The address of the selected range, that is the
//   active range, is written to the log.
function activeRangeFromSpreadsheetApp () {
 var activeRange = 
    SpreadsheetApp.getActiveRange();
 Logger.log(activeRange.getA1Notation());
}
// Get the active cell and print its containing
//  sheet name and address to the log.
// Try re-running after adding a new sheet
//  and selecting a cell at random.
function activeCellFromSheet () {
 var activeSpreadsheet = 
     SpreadsheetApp.getActiveSpreadsheet(),
     activeCell = 
       activeSpreadsheet.getActiveCell(),
     activeCellSheetName = 
       activeCell.getSheet().getSheetName(),
     activeCellAddress = 
      activeCell.getA1Notation();
 Logger.log('The active cell is in sheet: ' + 
              activeCellSheetName);
 Logger.log('The active cell address is: ' + 
              activeCellAddress);
}
// Print Range object properties 
// (all are methods) to log.
function printRangeMethods () {
 var rng = 
   SpreadsheetApp.getActiveRange();
 Logger.log(Object.keys(rng)
   .sort().join('\n'));
}
// Creating a Range object using two different 
//  overloaded versions of the Sheet 
//  "getRange()" method.
// "getSheets()[0]" gets the first sheet of the 
//   array of Sheet objects returned by 
//  "getSheets()".
// Both calls to "getRange()" return a Range
// object representing the same range address
//   (A1:B10).
function getRangeObject () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheets()[0],
      rngByAddress = sh.getRange('A1:B10'),
      rngByRowColNums = 
        sh.getRange(1, 1, 10, 2);
  Logger.log(rngByAddress.getA1Notation());
  Logger.log(
    rngByRowColNums.getA1Notation());
}
// Set a number of properties for a range.
// Add a new sheet.
// Set various properties for the cell 
//  A1 of the new sheet.
function setRangeA1Properties() {
  var ss = 
   SpreadsheetApp.getActiveSpreadsheet(),
      newSheet,
     rngA1;
 newSheet = ss.insertSheet();
 rngA1 = newSheet.getRange('A1');
 rngA1.setComment(
   'Hold The date returned by spreadsheet ' 
    + ' function "TODAY()"');
 rngA1.setFormula('=TODAY()');
 rngA1.setBackgroundColor('black');
 rngA1.setFontColor('white');
 rngA1.setFontWeight('bold');
}
// Demonstrate get methods for 'Range' 
//  properties.
// Assumes function "setRangeA1Properties()
//   has been run.
// Prints the properties to the log.
// Demo purposes only!
function printA1PropertiesToLog () {
  var rngA1 = 
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('RangeTest').getRange('A1');
  Logger.log(rngA1.getComment());
  Logger.log(rngA1.getFormula());
  Logger.log(rngA1.getBackground());
  Logger.log(rngA1.getFontColor());
  Logger.log(rngA1.getFontWeight());
}
// Starting with cell C10 of the active sheet,
// add comments to each of its adjoining cells
//  stating the row and column offsets needed
//  to reference the commented cell 
// from cell C10.
function rangeOffsetDemo () {
  var rng = 
    SpreadsheetApp.getActiveSheet()
     .getRange('C10');
  rng.setBackground('red');
  rng.setValue('Method offset()');
  rng.offset(-1,-1)
   .setComment('Offset -1, -1 from cell '
    + rng.getA1Notation());
  rng.offset(-1,0)
    .setComment('Offset -1, 0 from cell '
    + rng.getA1Notation());
  rng.offset(-1,1)
    .setComment('Offset -1, 1 from cell '
    + rng.getA1Notation());
  rng.offset(0,1)
  .setComment('Offset 0, 1 from cell '
  + rng.getA1Notation());
  rng.offset(1,0)
    .setComment('Offset 1, 0 from cell '
    + rng.getA1Notation());
  rng.offset(0,1)
    .setComment('Offset 0, 1 from cell '
    + rng.getA1Notation());
  rng.offset(1,1)
    .setComment('Offset 1, 1 from cell '
    + rng.getA1Notation());
  rng.offset(0,-1)
    .setComment('Offset 0, -1 from cell '
    + rng.getA1Notation());
  rng.offset(1,-1)
    .setComment('Offset -1, -1 from cell '
    + rng.getA1Notation());
}
// Passing a deliberately "bad" argument to the 
//  Range offset() method.
// The row offset argument is -1 but 
//  there is no row  above row 1
//   (cell A1's row).
// Google Apps Script gives error:
//   "It looks like someone else
// already deleted this cell."
function offsetError () {
  var rng = 
    SpreadsheetApp.getActiveSpreadsheet()
                    .getActiveSheet()
                    .getRange('A1');
  rng.offset(-1,0)
    .setValue('bad offset argument.');
}
function offsetError () {
  var rng = 
    SpreadsheetApp.getActiveSpreadsheet()
         .getActiveSheet().getRange('A1');
 Logger.log(rng.offset(-1,0).getValue());
}
// See the Sheet method getDataRange() in action.
// Print the range address of the used range for
//  a sheet to the log.
function getDataRange () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange();
  Logger.log(dataRange.getA1Notation());
}
// Read the entire data range of a sheet 
// into a JavaScript array.
// Uses the JavaScript Array.isArray()
//  method twice to verify that method
// getValues()returns an array-of-arrays. 
// Print the number of array elements 
// corresponding to the number of data 
//  range rows.
// Extract and print the first 10
//  elements of the array using the 
//  array slice() method.
function dataRangeToArray () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      dataRangeValues = dataRange.getValues();
  Logger.log(Array.isArray(dataRangeValues));
  Logger.log(Array.isArray(dataRangeValues[0]));
  Logger.log(dataRangeValues.length);
  Logger.log(dataRangeValues.slice(0, 10));
}
// Loop over the array returned by 
//  getRange() and a CSV-type output
// to the log using array join() method.
function loopOverArray () {
  var ss =
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      dataRangeValues = 
        dataRange.getValues(),
      i;
  for ( i = 0;
        i < dataRangeValues.length; 
        i += 1 ) {
    Logger.log(
      dataRangeValues[i].join(','));
  }
}
// In production code, this function would be 
//   re-factored into smaller functions.
// Read the data range into a JavaScript array.
// Remove and store the header line using the
//    array shift() method.
// Use the array filter() method with an anonymous
//   function as a callback to implement the 
//   filtering logic.
// Determine the element count of the 
//  filter() output array.
// Add a new sheet and store a reference to it.
// Create a Range object from the new
//   Sheet objectusing the getRange() method.
// The four arguments given to getRange() are:
//   (1) first column, (2) first row,
//   (3) last row, and (4) last column.  
// This creates a range corresponding to 
//  range address "A1:C5".
// Write the values of the filtered array to the 
//  newly created range.
function writeFilteredArrayToRange () {
  var ss = 
     SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      dataRangeValues = dataRange.getValues(),
      filteredArray,
      header = dataRangeValues.shift(),
      filteredArray,
      filteredArrayColCount = 3,
      filteredArrayRowCount,
      newSheet,
      outputRange;
  filteredArray = dataRangeValues.filter( 
    function (innerArray) { 
      if (innerArray[2] >= 40) {
        return innerArray;
      }});
  filteredArray.unshift(header);
  filteredArrayRowCount = filteredArray.length;
  newSheet = ss.insertSheet();
  outputRange = newSheet
                 .getRange(1, 
                           1, 
                           filteredArrayRowCount, 
                           filteredArrayColCount);
  outputRange.setValues(filteredArray);
}
function setRangeName () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet(),
      rng = sh.getRange('A1:B10'),
      rngName 'MyData';
  ss.setNamedRange(rngName, rng);
}
// Create a range object using the 
//  getDataRange() method.
// Pass the range and a colour name 
//  to function "setAlternateRowsColor()".
// Try changing the 'color' variable to
//   something like:
//   'red', 'green', 'yellow', 'gray', etc.
function call_setAlternateRowsColor  () {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange(),
      color = 'grey';
  setAlternateRowsColor(dataRange, color);
}

// Set every second row in a given range to
//   the given colour.
// Check for two arguments:
//   1: Range, 2: string for colour.
// Throw a type error if either argument
//  is missing or of the wrong type.
// Use the Range offset() to loop 
//   over the range rows.
// the for loop counter starts at 0.
// It is tested in each iteration with the
//   modulus operator (%).
// If i is an odd number, the if condition 
// evaluates to true and the colour 
//  change is applied.
// WARNING: If a non-existent colour is given, 
//  then the "color" is set to undefined
 (no color). NO error is thrown!
function setAlternateRowsColor (range, 
                                color) {
  if (range.toString() 
      !== 'Range') {
    throw {'name': 'TypeError',
           'message': 
           'The first argument to ' +
            '"setAlternateRowsColor()"  ' +
              ' must be type Range'};
  }
  if (typeof color !== 'string') {
    throw {'name': 'TypeError',
           'message': 
           'The second argument to ' +
            ' "setAlternateRowsColor()" ' +
            ' must be a string for a color,' +
            '  e.g. "red"'};
  }
  var i,
      startCell = range.getCell(1,1),
      columnCount = range.getLastColumn(),
      lastRow = range.getLastRow();
  for (i = 0; i < lastRow; i += 1) {
    if (i % 2) {
      startCell.offset(i, 0, 1, columnCount)
                 .setBackgroundColor(color);
    }
  } 
}
// Test function for 
//  "deleteLeadingTrailingSpaces()".
// Creates a Range object and passes
//   it to this function.
function call_deleteLeadingTrailingSpaces() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'english_premier_league',
      sh = ss.getSheetByName(sheetName),
      dataRange = sh.getDataRange();
  deleteLeadingTrailingSpaces(dataRange);
}

// Process each cell in the given range.
// If the cell is of type text 
//  (typeof === 'string') then
//  remove leading and trailing white space.  
//  Else ignore it.
// Code note: The Range getCell() method
//   takes two 1-based indexes 
//  (row and column).
//   This is in contrast to the offset() method. 
//  rng.getCell(0,0) will throw an error!
function deleteLeadingTrailingSpaces(range) {
  if (range.toString() !== 'Range') {
    throw {'name': 'TypeError',
           'message': 
           'Argument to ' + 
        '"deleteLeadingTrailingSpaces()" ' +
           'must be type Range'};
  }
  var i,
      j,
      lastCol = range.getLastColumn(),
      lastRow = range.getLastRow(),
      cell,
      cellValue;
  for (i = 1; i <= lastRow; i += 1) {
    for (j = 1; j <= lastCol; j += 1) {
      cell = range.getCell(i,j);
      cellValue = cell.getValue();
      if (typeof cellValue === 'string') {
        cellValue = cellValue.trim();
        cell.setValue(cellValue);
      }
    }
  } 
}
// Create a Sheet object for the active sheet.
// Pass the sheet object to 
//   "getAllDataRangeFormulas()"
// Create an array of the keys in the returned 
//  object in default "sort()".
// Loop over the array of sorted keys and 
//  extract the  values they keys map to.
// Write the output to the log.
function call_getAllDataRangeFormulas() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheet = ss.getActiveSheet(),
      sheetFormulas = getAllDataRangeFormulas(sheet),
      formulaLocations = 
        Object.keys(sheetFormulas).sort(),
      formulaCount = formulaLocations.length,
      i;
  for (i = 0; i < formulaCount; i += 1) {
    Logger.log(formulaLocations[i] + 
          ' contains ' + 
          sheetFormulas[formulaLocations[i]]);
  }
}
// Take a Sheet object as an argument, 
//  throw an error if the given argument 
//  is of the wrong type.
// Return an object literal where formula 
//  locations map to formulas for all formulas
//   in the input sheet data range.
// Loop through every cell in the data range.
// If a cell has a formula, 
//   store that cells location as 
//  the key and its formula as the value
//   in the object literal.
// Return the populated object literal.
function getAllDataRangeFormulas(sheet) {
  if (sheet.toString() !== 'Sheet') {
    throw {'name': 'TypeError',
           'message': 
           'Function "getAllDataRangeFormulas()" ' +
           ' expects a single argument of ' +
           ' type Sheet.'};
  }
  var dataRange = sheet.getDataRange(),
      i,
      j,
      lastCol = dataRange.getLastColumn(),
      lastRow = dataRange.getLastRow(),
      cell,
      cellFormula,
      formulasLocations = {},
      sheetName = sheet.getSheetName(),
      cellAddress;
  for (i = 1; i <= lastRow; i += 1) {
    for (j = 1; j <= lastCol; j += 1) {
      cell = dataRange.getCell(i,j);
      cellFormula = cell.getFormula();
      if  (cellFormula) {
        cellAddress = sheetName + '!' + 
            cell.getA1Notation();
        formulasLocations[cellAddress] =
          cellFormula;
      }
    }
  }
  return formulasLocations;
}
// Call copyColumns() function passing it:
//  1: The active sheet
//  2: A newly inserted sheet
//  3: An array of column indexes to copy
//     to the new sheet
// The output in the newly inserted sheet 
//  contains the columns for the indexes
//   given in the array in the 
// order specified in the array.
function call_copyColumns() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      inputSheet = ss.getActiveSheet(),
      outputSheet = ss.insertSheet(),
      columnIndexes = [4,3,2,1];
  copyColumns(inputSheet, 
              outputSheet, 
              columnIndexes);
}

// Given an input sheet, an output sheet,
// and an array:
// Use the numeric column indexes in 
//  the array to copy those columns from 
//  the input sheet to the output sheet.
// The function checks its input arguments
//   and throws an error
// if they are not Sheet, Sheet, Array.
// The array is expected to be an array of
//  integers but it does 
//   not check the array element types
function copyColumns(inputSheet,
                     outputSheet,
                     columnIndexes) {
  if (! (inputSheet.toString() ===
         'Sheet'
         && 
         outputSheet.toString() === 'Sheet') ) {
    throw {'name': 'TypeError',
           'message': 'Function ' + '
           "copyColumns()": ' + 
           'First two arguments must ' +
           ' be Sheet objects'};
  }
  if (! Array.isArray(columnIndexes)) {
    throw {'name': 'TypeError',
           'message': 'Function ' + 
           '"copyColumns()": ' +
           'Third argument has to be ' +
           ' an array of indexes'};
  }
  var dataRangeRowCount = 
       inputSheet.getDataRange()
                 .getNumRows(),
      columnsToCopyCount = 
         columnIndexes.length,
      i,
      columnIndexesCount,
      valuesToCopy = [];
  for (i = 0; 
       i < columnsToCopyCount;
       i += 1) {
    valuesToCopy = 
      inputSheet
        .getRange(1, 
                  columnIndexes[i], 
                  dataRangeRowCount,
                  1).getValues();
    outputSheet
     .getRange(1, 
      i+1, 
      dataRangeRowCount, 
      1).setValues(valuesToCopy);
  } 
}

// Chapter 6

// Test connection to a  MySQ
//  cloud instance created earlier.
// Check log for output.
function connectMySqlCloud() {
  var connStr = 
      'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      conn;
  try {
    conn = 
      Jdbc.getCloudSqlConnection(connStr);
    Logger.log('Connection OK!');
  } catch (err) {
    Logger.log(err);
    throw err;
  } finally {
    if (conn) {
      conn.close();
    }
  }
}
// Execute a CREATE TABLE DDL statement for a 
//  database named "contacts".
function createTable() {
  var connStr = 
      'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      conn,
      stmt,
      ddl;
  ddl = 'CREATE TABLE person(' +
         '  person_id  MEDIUMINT ' +
            'AUTO_INCREMENT' +
            ' NOT NULL PRIMARY KEY,' +
         '  first_name VARCHAR(100) NOT NULL,' +
         '  last_name VARCHAR(100) NOT NULL,' +
         '  date_of_birth DATE,' +
         '  height_cm SMALLINT)';
  try {
    conn = 
      Jdbc.getCloudSqlConnection(connStr);
    stmt = conn.createStatement();
    stmt.execute(ddl);
    Logger.log('Table created!');
  } catch (ex) {
    Logger.log(ex);
    throw(ex);
  } finally {
    Logger.log('Cleaning up.');
    if (stmt) {
      stmt.close();
    }
    if (conn) {
      conn.close();
    }
  }  
}
// Add 6 rows to newly created table.
// Data source is a JavaScript array-of-arrays.
// Uses bind parameters.
// Executes an SQL INSERT INTO statement 
//  within a for loop.
function addRowsToTable() {
  var connStr = 'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      conn,
      dml,
      prepStmt,
      rows,
      i,
      row,
      firstName,
      lastName,
      dateOfBirth,
      heightcm;
  dml = 'INSERT INTO person(first_name, ' +
                            'last_name, ' +
                            'date_of_birth, ' +
                            'height_cm) ' +
          'VALUES(?, ?, ?, ?)';
  rows = [['Joe', 'Grey', '1970-06-11', 182],
          ['Raj', 'Patel', '1975-03-13', 188],
          ['Amy', 'Lopez', '1972-08-17', 166],
          ['Bill', 'Grimes', '1954-10-20', 181],
          ['Jane', 'Rice', '1961-04-30', 170],
          ['Alex', 'Lee', '1982-08-06', 190]
         ];
  try {
    conn = 
      Jdbc.getCloudSqlConnection(connStr);
    prepStmt = conn.prepareStatement(dml);
    for (i = 0; i < rows.length; i += 1) {
      row = rows[i];
      firstName = row[0];
      lastName = row[1];
      dateOfBirth = row[2];
      heightcm = row[3];
      prepStmt.setString(1, firstName);
      prepStmt.setString(2, lastName);
      prepStmt.setString(3, dateOfBirth);
      prepStmt.setInt(4, heightcm);
      prepStmt.execute();
    }
    Logger.log('Loaded row count: ' + i);
  } catch(ex) {
    Logger.log(ex);
    throw(ex);    
  } finally {
    if (prepStmt) {
      prepStmt.close();
    }
    if (conn) {
      conn.close();
    }
  }
}
// Retrieve all rows from a table and 
//  write the to a newly added sheet.
// If re-running this code, ensure 
//  that added sheet is deleted 
//  or re-named.
function writeDatabasRowsToSpreadsheet() {
  var connStr = 'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      conn,
      sql = 'SELECT * FROM person',
      stmt,
      rs,
      colCount,
      colVal,
      rowVals = [],
      i,
      ss = SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.insertSheet();
  sh.setName('QueryResults');
  try {
    conn = 
      Jdbc.getCloudSqlConnection(connStr);
    stmt = conn.createStatement();
    rs = stmt.executeQuery(sql);
    colCount = rs.getMetaData()
                 .getColumnCount();
    Logger.log('Col count: ' + colCount);
    while (rs.next()) {
      for (i = 1; i <= colCount; i += 1) {
        colVal = rs.getString(i);
        rowVals.push(colVal);
      }
      sh.appendRow(rowVals);
      rowVals = [];
    }
  } catch(ex) {
    Logger.log(ex);
    throw(ex);    
  } finally {
    if (rs) {
      rs.close();
    }
    if (stmt) {
      stmt.close()
    };
    if (conn) {
      conn.close();
    }
  }
}
// Execute the MySQL "SHOW TABLES"
// statement and print the table
//  names in the target database
//  to the log.
// Exception handling is missing!
function showTables() {
  var connStr = 
      'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      conn,
      stmt,
      rs,
      sql;
  sql = 'SHOW TABLES';
  conn = 
    Jdbc.getCloudSqlConnection(connStr);
  stmt = conn.createStatement();
  rs = stmt.executeQuery(sql);
  while (rs.next()) {
    Logger.log(rs.getString(1));
  }
  rs.close();
  stmt.close();
  conn.close();
}
// Print some metadata about the database created
//  in earlier examples.
// Adds a sheet named "databaseMetadata".
// If re-running, remove sheet added previously.
// Exception handling dropped.
function writeMySQLMetadataToSheet() {
  var connStr = 'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      dbMetadata,
      rsTables,
      rsColumns,
      tableNames = [],
      i,
      tableCount,
      ss,
      sh;
    ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
    sh = ss.insertSheet()
    conn = 
      Jdbc.getCloudSqlConnection(connStr);
    dbMetadata = conn.getMetaData();
    sh.setName('DatabaseMetadata');
    sh.appendRow(['Major Version', 
        dbMetadata.getDatabaseMajorVersion()]);
    sh.appendRow(['Minor Version',
        dbMetadata.getDatabaseMinorVersion()]);
    sh.appendRow(['Product Name', 
        dbMetadata.getDatabaseProductName()]);
    sh.appendRow(['Product Version',
        dbMetadata.getDatabaseProductVersion()]);
    sh.appendRow(['Supports transactions',
        dbMetadata.supportsTransactions()]);
    rsTables = dbMetadata.getTables(
        null, null, null, ['TABLE']);
    while (rsTables.next()) {
      tableNames.push(rsTables.getString(3));
    }
    tableCount = tableNames.length;
    sh.appendRow(
      ['Table Names And Columns Names Are:']);
    for (i = 0; i < tableCount; i += 1) {
      rsColumns = dbMetadata.getColumns(
        null, null, tableNames[i], null);
      while (rsColumns.next()) {
        sh.appendRow(
     [tableNames[i], rsColumns.getString(4)]);
      }
    }
    rsTables.close();
    rsColumns.close();
    conn.close();
}
/*
-- SQL for GAS that follows:
SELECT
  tab.table_name,
  tab.engine,
  tab.table_rows,
  col.column_name,
  col.column_type
FROM
  INFORMATION_SCHEMA.TABLES AS tab
  JOIN INFORMATION_SCHEMA.COLUMNS AS col
    ON tab.table_name = col.table_name
WHERE
  tab.table_schema = ?
  AND
  	tab.table_type = 'BASE TABLE';
*/
// Driver function for "getDocText()"
function test_getDocText() {
  // Replace <docID> with your doc ID.
  var docID = <docID>;
  Logger.log(getDocText(docID));
}

// Given a Document ID return
// its text as a JavaScript string.
// Requires "Authorization" to run.
function getDocText(docID) {
  var txt;
  try {
    doc = DocumentApp.openById(docID);
    txt = doc.getText();
    return txt;
  } catch (ex) {
    Logger.log('Error in "getSql()": ' + ex);
    throw(ex);
  }
}
// Read in an SQL file as text.
// Connect to the MySQL instance containing the 
//  database of interest.
// Pass in a database name to bind parameter and 
//   generate a RecordSet.
// Add a sheet to the active spreadsheet with a 
//  user-specified name
// Process the RecordSet and write the output to
//   the newly added sheet.
// Lacks argument checking and
//  exception handling!
function writeMySqlDBSummaryToSheet(sqlFileId, 
                                    dbName,
                                    sheetName) {
  var connStr = 'jdbc:google:rdbms://' +
       'elwarbito:chapter6/contacts',
      conn,
      sql,
      prepStmt,
      recSet,
      ss,
      sh,
      colCount,
      colNames = [],
      i,
      rowVals = [];
    ss = SpreadsheetApp.getActiveSpreadsheet();
    sh = ss.insertSheet();
    sh.setName(sheetName);
    // Read in the SQL file contents to a string.
    sql = getDocText(sqlFileId);
    conn = 
      Jdbc.getCloudSqlConnection(connStr);
    prepStmt = conn.prepareStatement(sql);
    // Set the bind parameter to the database name.
    prepStmt.setString(1, dbName);
    recSet = prepStmt.executeQuery();
    colCount = recSet.getMetaData()
                     .getColumnCount();
    // Get Column Names.
    for (i = 1; i <= colCount; i += 1) {
      colNames.push(recSet.getMetaData()
                          .getColumnName(i))
    }
    // Write column names to first row of new sheet.
    sh.appendRow(colNames);
    // Get column values from RecordSet.
    while (recSet.next()) {
      for (i = 1; i <= colCount; i += 1) {
        rowVals.push(recSet.getString(i));
      }
      sh.appendRow(rowVals);
      rowVals = [];
    }
    prepStmt.close();
    conn.close();
}
// Driver function for
//  "writeMySqlDBSummaryToSheet()"
function test_writeMySqlDBSummaryToSheet() {
  // Replace <docID> with your doc ID.
  var sqlFileId = '<doc_id>',
      dbName = 'contacts',
      sheetName = 'DBSummary';
  writeMySqlDBSummaryToSheet(sqlFileId, 
                             dbName,
                             sheetName);
}

// Chapter 7

// Pre-defined function that runs as a
//  "trigger" when the spreadsheet is opened.
// Here it is set to build the menu.
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      menuItems;
  menuItems = [{name: 'Calculate BMI', 
               functionName: 'calculateBmi'}];
 ss.addMenu('Custom Menu', menuItems);
}
// Displays two input box prompts:
//  First for weight in kilograms
//  Second for height in meters
// Displays result of calculation in a
//   message box.
function calculateBmi() {
  var bmiTitle = 'BMI Calculation',
      massPrompt = 'Enter weight in Kg:',
      bmiButtons = Browser.Buttons.OK_CANCEL,
      inputMass,
      heightPrompt = 'Enter height in meters:',
      inputHeight,
      msgTitle = 'BMI Calculation',
      msgPrompt = 'BMI is: ',
      bmiResult;
  try {
    inputMass = Browser.inputBox(bmiTitle, 
                                 massPrompt, 
                                 bmiButtons);
    if (inputMass === 'cancel') {
      return;
    }
    inputHeight = Browser.inputBox(bmiTitle, 
                                   heightPrompt, 
                                   bmiButtons);
    if (inputHeight === 'cancel') {
      return;
    }
    bmiResult = getBmiMetric(
                  parseFloat(inputMass, 10),
                  parseFloat(inputHeight, 10));
    Browser.msgBox(Number(bmiResult).toFixed(2));
  }
  catch (ex) {
    Browser.msgBox('Error: ' + 
                   ex.name + ': ' + 
                   ex.message);
  }
}
// Performs Body-Mass Index (BMI) calculation.
// Requires two numeric arguments:
//  1: Weight in kilograms
//  2: Height in meters
// Can be called as a spreadsheet function:
//   cell formula example 
//   "=getBmiMetric(90, 1.90)"
function getBmiMetric(massKg, heightM) {
  Logger.log(massKg);
  if (! (typeof massKg === 'number' 
      && typeof heightM === 'number')) {
    throw {'name': 'TypeError',
           'message': 'Function requires ' +
              'weight in Kg and ' +
              'height in meters'};
  }
  return (massKg/(heightM * heightM));
}
// 7.3.1 First Form
// Menu is created and added when the containing
//  spreadsheet is opened.
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      menuItems;
  menuItems = [{name: 'Show Form', 
               functionName: 'guiExample1'}];
 ss.addMenu('Custom Menu', menuItems);
}
// Close the GUI
function close() {
	var ui = UiApp.getActiveApplication();
  return ui.close();  
}
// Get the active user name.
// Write the active user name as text in a label
//  widget on the form.
function displayUser() {
  var ui = UiApp.getActiveApplication(),
      user = Session.getActiveUser(),
      lblUser = ui.getElementById('lblUser_ID');
  lblUser.setText(user);
  return ui;
}
// Create a simple GUI form with
//   two buttons and one label.
// Uses default sizes and layouts
function guiExample1() {
  var ui = UiApp.createApplication(),
      ss = SpreadsheetApp.getActiveSpreadsheet(),
      lblUser = ui.createLabel(''),
      btnUser = ui.createButton('User'),
      btnClose = ui.createButton('Close'),
      closeFunction = 'close',
      closeHandler = 
        ui.createServerHandler(closeFunction),
      userFunction = 'displayUser',
      userHandler = 
        ui.createServerHandler(userFunction);
  lblUser.setId('lblUser_ID');
  ui.setTitle('Display User ID');
  btnUser.addClickHandler(userHandler);
  btnClose.addClickHandler(closeHandler);
  ui.add(lblUser);
  ui.add(btnUser);
  ui.add(btnClose);
  ss.show(ui);  
}
// Simple data entry form GUI.
// Contains paired labels and text boxes
// Uses panels as widget containers to
//   get desired layout.
// To run add use the "onOpen()" trigger or
//   execute it from the Script Editor.
function guiExample2() {
	var ui = UiApp.createApplication(),
		ss = SpreadsheetApp.getActiveSpreadsheet(),
        pnlMain = ui.createVerticalPanel(),
        pnlName = ui.createHorizontalPanel(),
        pnlEmail = ui.createHorizontalPanel(),
        pnlBtns = ui.createHorizontalPanel(),
        lblName = ui.createLabel('Name: '),
        txtName = ui.createTextBox(),
        lblEmail = ui.createLabel('Email: '),
        txtEmail = ui.createTextBox(),
        btnAddRecord = 
          ui.createButton('Add Record'),
        btnClose = ui.createButton('Close');
  ui.setTitle('Data Entry Form');
  ui.setHeight(80);
  ui.setWidth(200);
  btnClose.setWidth(90);
  btnAddRecord.setWidth(90);
  txtName.setWidth(150);
  txtEmail.setWidth(150);
  ui.add(pnlMain);
  pnlName.add(lblName);
  pnlName.add(txtName);
  pnlMain.add(pnlName);
  pnlEmail.add(lblEmail);
  pnlEmail.add(txtEmail);
  pnlMain.add(pnlEmail);
  pnlBtns.add(btnAddRecord);
  pnlBtns.add(btnClose);
  pnlMain.add(pnlBtns);
  ss.show(ui);
}
// Close the current UI
function close() {
    var ui = UiApp.getActiveApplication();
    return ui.close();
}

// Callback action for button labelled "Add Row"
// Take the values from two text boxes in the 
//  active GUI and add them as rows to the 
//  active sheet.
// Prepare for next input by:
// re-setting the text boxes to empty strings
// Set the focus to the first text box
function addRow(e) {
    var ui = UiApp.getActiveApplication(),
        sheet = 
          SpreadsheetApp.getActiveSheet(),
        name = e.parameter.txtName_Name,
        email = e.parameter.txtEmail_Name,
        txtName =
          ui.getElementById('txtName_Id'),
        txtEmail = 
          ui.getElementById('txtEmail_Id'),
        activeUser = Session.getActiveUser();
    sheet.appendRow([name, email, activeUser]);
    txtName.setValue('');
    txtEmail.setValue('');
    txtName.setFocus(true);
    return ui;
}

// Build a GUI with two labels, two text 
//  boxes and two buttons.
// Layout uses panels as in previous example.
function dataEntryForm() {
    var ui = UiApp.createApplication(),
        ss = 
          SpreadsheetApp.getActiveSpreadsheet(),
        uiTitle = 'Data Entry Form',
        pnlMain = ui.createVerticalPanel(),
        pnlName = ui.createHorizontalPanel(),
        pnlEmail = ui.createHorizontalPanel(),
        pnlButtons = ui.createHorizontalPanel(),
        lblName = ui.createLabel('Name:'),
        lblEmail = ui.createLabel('Email:'),
        txtName = ui.createTextBox(),
        txtEmail = ui.createTextBox(),
        btnAddRow =
          ui.createButton('Add Record'),
        btnExit = ui.createButton('Close'),
        exitHandler = 
          ui.createServerHandler('close'),
        addRowHandler = 
          ui.createServerHandler('addRow');
    pnlName.add(lblName);
    pnlName.add(txtName);
    pnlEmail.add(lblEmail);
    pnlEmail.add(txtEmail);
    pnlMain.add(pnlName);
    pnlMain.add(pnlEmail);
    pnlButtons.add(btnAddRow);
    pnlButtons.add(btnExit);
    pnlMain.add(pnlButtons);
    ui.add(pnlMain);
    ui.setWidth(200);
    ui.setHeight(150);
    btnExit.setWidth(80);
    btnExit.addClickHandler(exitHandler);
    btnAddRow.setWidth(80);
    btnAddRow.addClickHandler(addRowHandler);
    addRowHandler.addCallbackElement(txtName);
    addRowHandler.addCallbackElement(txtEmail);
    txtName.setName('txtName_Name');
    txtEmail.setName('txtEmail_Name');
    txtName.setId('txtName_Id');
    txtEmail.setId('txtEmail_Id');
    ui.setTitle(uiTitle);
    txtName.setFocus(true);
    ss.show(ui);
}
// Given a SpreadsheetApp instance 
//  as an argument:
//  Return a list of sheet names.
function getSheetNames(ss) {
  var sheets = ss.getSheets(),
      sheetNames = [],
      i;
  for (i = 0; i < sheets.length; i += 1) {
      sheetNames.push(sheets[i].getSheetName());
  }
  return sheetNames;
}
// Executes when the listbox selection 
//  is changed.
// Extracts the address of the selected
//   sheet name
// data range (used range) and writes
//  the address to a text box.
function writeSheetInfoToGui(e) {
    var ui = UiApp.getActiveApplication(),
        ss = 
          SpreadsheetApp.getActiveSpreadsheet(),
        selectedSheetName = 
          e.parameter.lbxSheetNames_Name,
        sheet,
        dataRngAddress,
        txtDataRng = 
          ui.getElementById('txtDataRng_Id');
    // Avoid error if empty place-holder 
    //  is selected.
    if (selectedSheetName.length > 0) {
      sheet = 
        ss.getSheetByName(selectedSheetName);
      dataRngAddress = 
        sheet.getDataRange().getA1Notation();
      txtDataRng.setValue(dataRngAddress);
    } else {
      txtDataRng.setValue('');
    }
    return ui; // NB always remember this line!
}
// Build a GUI to demonstrate list boxes.
// Sheet names of the active spreadsheet
//   are added to the list box.
// When the selected sheet name in the
//   list box is changed, the selected 
//  sheet data range is written to the 
// text box.
// Shows how an application can respond 
//  dynamically to user interaction.
function spreadsheetInfo() {
    var ui = UiApp.createApplication(),
        ss = 
          SpreadsheetApp.getActiveSpreadsheet(),
        uiTitle = 'Spreadsheet Info',
        pnlMain = ui.createVerticalPanel(),
        pnlSheetNames = 
         ui.createHorizontalPanel(),
        pnlDataRng = ui.createHorizontalPanel(),
        pnlRowCount = ui.createHorizontalPanel(),
        pnlColCount = ui.createHorizontalPanel(),
        lblDataRng = 
          ui.createLabel('Data Range Address'),
        txtDataRng = ui.createTextBox(),    
        lblSheetNames =
          ui.createLabel('Sheet Names:'),
        lbxSheetNames = ui.createListBox(),
        sheetNames = getSheetNames(ss),
        i,
        lbxHandler = 
      ui.createServerHandler(
                 'writeSheetInfoToGui');
    lbxSheetNames.addChangeHandler(lbxHandler);
    ui.add(pnlMain);
    // Add an emty item to the list box
    lbxSheetNames.addItem('');
    // Add the sheet names to the list box
    for (i = 0; i <sheetNames.length; i += 1) {
        lbxSheetNames.addItem(sheetNames[i]);
    }
    lbxSheetNames.setName('lbxSheetNames_Name');
    txtDataRng.setId('txtDataRng_Id');
    // Sets spacing between contained panels.
    pnlMain.setSpacing(10);
    pnlSheetNames.add(lblSheetNames);
    pnlSheetNames.add(lbxSheetNames);
    pnlDataRng.add(lblDataRng);
    pnlDataRng.add(txtDataRng);
    pnlMain.add(pnlSheetNames);
    pnlMain.add(pnlDataRng);
    lbxSheetNames.setVisibleItemCount(3);
    lblSheetNames.setWidth(120);
    lbxSheetNames.setWidth(120);
    lblDataRng.setWidth(120);
    ui.setWidth(300);
    ui.setHeight(150);
    ui.setTitle(uiTitle);
    ss.show(ui);
}


// Chapter 8

// Write the names of folders in the user's
// Google Drive to the logger
function listDriveFolders() {
  var folders = DocsList.getAllFolders(),
      i;
  if (folders.length === 0) {
    Logger.log("No folders in this Drive");
    return;
  }
  Logger.log("Folder Names:");
  for (i = 0; i< folders.length; i += 1) {
    Logger.log(folders[i].getName());
  } 
}
// Write the names of files in the user's
// Google Drive to the logger.
function listDriveFiles() {
  var files = DocsList.getAllFiles(),
      i;
  if (files.length === 0) {
    Logger.log("No files found in this Drive");
    return;
  }
  Logger.log("File Names:");
  for (i = 0; i < files.length; i += 1) {
    Logger.log(files[i].getName());
  }
}
// Display the name of the top-level
// Google Drive folder.
function showRoot() {
  var root = DocsList.getRootFolder();
  Browser.msgBox(root.getName());
}
// Demonstration code only to show how 
// Google Drive allows duplicate file
// names and duplicate folder names
// within the same folder.
function makeDuplicateFilesAndFolders() {
  SpreadsheetApp.create('duplicate spreadsheet');
  SpreadsheetApp.create('duplicate spreadsheet');
  DocsList.createFolder('duplicate folder');
  DocsList.createFolder('duplicate folder');
}
// Demonstrates how folders can have the 
//  same name and parent folder 
//  (Root in this instance) but yet
// have different IDs.
function writeFolderNamesAndIds() {
  var folders = DocsList.getFolders(); 
  folders.forEach(
    function (folder) { 
      Logger.log(folder.getName() + 
      + ': ' + folder.getId());
   });
}
// Remove folders and files that are 
// identified by name from the 
// root folder.
// Caution when using names to
//  identify objects to delete!
// Ensure the name used to identify
//  objects to delete is not used
//  by objects to be retained.
function removeFoldersAndFiles() {
  var root = DocsList.getRootFolder(),
      foldersToRemove  = [],
      filesToRemove = [];
  foldersToRemove = 
    root.getFolders().filter(
      function (folder) { 
        return (folder.getName() 
                 === 
               'duplicate folder'); 
      });
  foldersToRemove.forEach(
    function (folder) { 
      folder.setTrashed(true);
    });
  filesToRemove = 
    root.getFiles().filter(
      function (file) { 
        return (file.getName()
                  === 
               'duplicate spreadsheet');
      });
  filesToRemove.forEach(
    function (file) { 
      file.setTrashed(true);
    });
}
// Create a test folder and a test file.
// Add the test file to the test folder.
function addNewFileToNewFolder() {
  var newFolder = DocsList.createFolder('Test Folder'),
      newSpreadsheet = SpreadsheetApp.create('Test File'),
      newSpreadsheetId = newSpreadsheet.getId(),
      newFile = DocsList.getFileById(newSpreadsheetId);
  newFile.addToFolder(newFolder);
}
// Remove the file from root folder.
// Does not delete it!
// ID was taken from the URL.
function removeTestFileFromRootFolder() {
  var root = DocsList.getRootFolder(),
      id = <fileID>,
      file = DocsList.getFileById(id);
  file.removeFromFolder(root);
}
// Remove a folder and its file contents
function deleteFolder() {
  var root = DocsList.getRootFolder(),
      folders = root.getFolders(),
      i;
  for (i = 0; i < folders.length; i += 1) {
    if (folders[i].getName() === 'Test Folder') {
      folders[i].setTrashed(true);
    }
  }
}
// Write properties of the Folder and 
// File objects to separate newly-created
//  sheets.
// Create a test folder and a test Google 
//  Spreadsheet.
// For the Spreadsheet, we want to treat it as
//  a generic file for the purpose of retrieving
//  File object  properties.
// Call the function 
//  "writeObjectPropertiesToSheet()"
// twice passing it the Folder and File objects
//  in turn with names for the sheets to where
// the object properties are written.
// Since this function creates a folder 
//  and a file,it requires authorisation.
function writePropertiesToSheets() {
  var newFolder = 
        DocsList.createFolder("Test Folder"),
      newSpreadsheet =
        SpreadsheetApp.create('Test File'),
      newSpreadsheetId =
        newSpreadsheet.getId(),
      newFile =
        DocsList.getFileById(newSpreadsheetId),
      folderPropsSheetName = 'FolderProperties',
      filePropsSheetName = 'FileProperties';
  writeObjectPropertiesToSheet(newFolder, 
                               folderPropsSheetName);
  writeObjectPropertiesToSheet(
                       newFile, 
                       filePropsSheetName);
}
// Given any object and a string for a sheet 
//  name, insert a new sheet with the given
//  name and extract all the properties from
//  the object as a sorted array.
// Loop over the array and write the property 
//  names to the new sheet.
// Function checks for two arguments:
//  1: any type of object
//  2: A string for the new sheet name
function writeObjectPropertiesToSheet(obj, 
                                   sheetName) {
  if (!(typeof obj === 'object' && 
        typeof sheetName === 'string')) {
    throw {'name': 'TypeError',
           'message': 'Function expects ' +
                  'an object and a string.'};
  }
  var sh = 
      SpreadsheetApp.getActiveSpreadsheet()
                         .insertSheet(),
      objProps = Object.keys(obj).sort(),
      i;
  sh.setName(sheetName);
  for (i = 0; i < objProps.length; i += 1) {
    sh.appendRow([objProps[i]]);
  }
}
// Creates two files in "My Drive".
// (1) A spreadsheet and (2) A document.
// Records the name, ID and URL for the
// two new files in the active spreadsheet.
// Gives a range name to the two data rows.
// Script requires authorisation.
function createFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getActiveSheet(),
      newSsName = 'myspreadsheet',
      newSs = SpreadsheetApp.create(newSsName),
      newDocName = 'mydocument',
      newDoc = DocumentApp.create(newDocName);
  sh.appendRow(['File Name', 
                'File ID', 
                'File URL']);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  sh.appendRow([newSsName, 
                newSs.getId(), 
                newSs.getUrl()]);
  sh.appendRow([newDocName, 
                newDoc.getId(), 
                newDoc.getUrl()]);
  ss.setNamedRange('FileDetails',
                   sh.getRange(2, 1, 2, 3));
}
// Adds a viewer to the document created above.
// Uses the range name to access the file ID.
// Requires authorisation.
function addViewer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      rngName = 'FileDetails',
      inputRng = ss.getRangeByName(rngName),
      docId = inputRng.getCell(2, 2).getValue(),
      doc = DocsList.getFileById(docId),
      newViewer =
  'mick@javascript-spreadsheet-programming.com';
  doc.addViewer(newViewer);
}
// Adds an editor to the spreadshet whose ID
// is stored first row and second column of the
// named range
function addEditor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      rngName = 'FileDetails',
      inputRng = ss.getRangeByName(rngName),
      ssToShareId = 
        inputRng.getCell(1, 2).getValue(),
      ssToShare = 
        DocsList.getFileById(ssToShareId),
      newEditor =
'mick@javascript-spreadsheet-programming.com';
  ssToShare.addEditor(newEditor);
}
// Create a new folder.
// Add the spreadsheet and document created
//  earlier to this new folder.
// As in earlier examples, the file objects
//  are retrieved using their IDs that
//  were stored in the active spreadsheet.
// Remove them from the root folder.
function putFilesToShareInNewFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      rootFolder = DocsList.getRootFolder(),
      newFolder = 
        DocsList.createFolder('ToShareWithMick'),
      rngName = 'FileDetails',
      inputRng = ss.getRangeByName(rngName),
      ssToShareId =
       inputRng.getCell(1, 2).getValue(),
      ssToShare =
        DocsList.getFileById(ssToShareId),
      docId = inputRng.getCell(2, 2).getValue(),
      doc = DocsList.getFileById(docId);
  ssToShare.addToFolder(newFolder);
  doc.addToFolder(newFolder);
  ssToShare.removeFromFolder(rootFolder);
  doc.removeFromFolder(rootFolder);
}
// Add an editor to the new folder.
// The new folder is identified by its name.
//  All folders are checked for the name 
//  (filter) and the first one with 
//  indicated name is returned.
// The folder ID is displayed in a messagebox
//  for verification purposes only.
// An editor is added to the Folder object.
function shareFolderForEditing() {
  var folders = 
       DocsList.getRootFolder().getFolders(),
      folderName = 'ToShareWithMick',
      folderToShare,
      folderId,
      collaborator = 
  'mick@javascript-spreadsheet-programming.com';
  folderToShare = folders.filter(
    function (folder) {
      return (folder.getName() 
                === 
              folderName);
    })[0];
  folderId = folderToShare.getId();
  Browser.msgBox(folderId);
  folderToShare.addEditor(collaborator);
}
// Given a Folder object and a sub-folder name,
// check to see if the sub-folder name exists
// in the given parent folder.
// Return true if it does, else return false
// Checks its arguments for a Folder object 
//  and a string.
function folderNameExists(parentFolder, 
                          subfolderName) {
  if (! (parentFolder.toString() ===
          'Folder' && 
         typeof subfolderName ===
         'string') ) {
    throw {'name': 'TypeError', 
           'message': 'Requires two ' + 
           'arguments: ' +
           'Folder object and ' +
           'String for folder name'};
  }
  var folderNames;
  folderNames = parentFolder
    .getFolders().map(
    function  (folder) {
      return folder.getName();
    });
  Logger.log(folderNames);
  if (folderNames
      .indexOf(subfolderName) > -1) {
    return true;
  } else {
    return false;
  }   
}
// Code to test "folderNameExists()"
// Change the value of folderName to existing
//  and non-existent folder names to see
// how the tested function operates!
function test_folderNameExists() {
  var folder = DocsList.getRootFolder(),
      folderName = 'Javascript';
  if (folderNameExists(folder, folderName)) {
    Logger.log("yup");
  } else {
    Logger.log("Nope");
  }
}
// Return an array of File objects
// for all files in the script user's
// Google Drive
function getAllFiles() {
  return DocsList.getAllFiles();
}

// Filter all files based on a given cut-off date
//  and return an array of files older than the
//  cut-off date.
function getFilesOlderThan(cutoffDate) {
  var filesOlderThan = 
        getAllFiles().filter(
        function(file) {
          return (file.getDateCreated() 
                   < 
                  cutoffDate)
        });
  return filesOlderThan;
}

// Write some file details for files older than
//  a specified date to a new sheet.
function test_getFilesOlderThan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      shName = 'OldFiles',
      sh = ss.insertSheet(),
      // 'July 28, 2010', months are 0-11 in JS
      testDate = new Date(2010, 6, 28);
      oldFiles = getFilesOlderThan(testDate);
  sh.setName(shName);
  oldFiles.forEach(
    function (file) {
      sh.appendRow([file.getName(), 
                    file.getSize(), 
                  file.getDateCreated()]);
    });
}
// Given a file object, return number
// of days since its creation.
// Checks argument is type File.
function getFileAgeInDays(file) {
  if (! file.toString() === 'File' ) {
    throw {'name': 'TypeError',
           'message': 'Argument must be ' +
           'type File'};
  }
  var today = new Date(),
      createdDate = file.getDateCreated(),
      msecPerDay = 1000 * 60 * 60 * 24,
      fileAgeInDays =
     (today - createdDate)/msecPerDay;
  return Math.round(
    fileAgeInDays).toFixed(0);
}
// Get a test file ID from the URL
// and assign it to variable fileID
// Run and check the log
function test_getFileAgeInDays() {
  var fileId = <FileID>,
      file = 
        DocsList.getFileById(fileId),
      fileAgeInDays =
       getFileAgeInDays(file);
  Logger.log('File ' + 
             file.getName() + 
             ' is ' + 
             fileAgeInDays + 
             ' days old.');
}
// Return an array of empty folder objects 
// in the user's Google drive.
// An empty folder is one with no associated 
//  files or sub-folders.
function getEmptyFolders() {
  var folders = DocsList.getAllFolders(),
      emptyFolders = [];
  emptyFolders =
    folders.filter( 
      function (folder) {
        return (
        folder.getFiles().length === 0 
          &&
        folder.getFolders().length === 0);
    });
  return emptyFolders;
}

// Write the IDs and names of
//  all empty folders to the 
//  log.
function processEmptyFolders() {
  var emptyFolders =
      getEmptyFolders();
  emptyFolders.forEach(
    function (folder) {
      Logger.log(folder.getName() +
                 ': ' + 
                 folder.getId());
    });
}
// Return object mapping each
//  file name to an array of file IDs.
// If file names are not duplicated,
//  the arrays they map to will
// have a single element (file ID).
function getFilenameIdMap() {
  var files = DocsList.getAllFiles(),
      filenameIdMap = {},
      i,
      filename,
      fileId;
  for (i = 0; i < files.length; i += 1) {
    filename = files[i].getName();
    fileId = files[i].getId();
    if (filename in filenameIdMap) {
      filenameIdMap[filename]
        .push(fileId);
    } else {
      filenameIdMap[filename] = [];
      filenameIdMap[filename]
        .push(fileId);
    }
  }
  return filenameIdMap;
}
// Return an array of file IDs for files
// with duplicate names.
// Loops over the object returned by
// getFilenameIdMap() and returns
// only those file IDs for duplicate
// file names.
function getDuplicateFilenameIds() {
  var filenameIdMap = getFilenameIdMap(),
      filename,
      duplicateFilenameIds = [];
  for (filename in filenameIdMap) {
    if (filenameIdMap[filename].length > 1) {
      duplicateFilenameIds =
        duplicateFilenameIds
          .concat(filenameIdMap[filename]);
    }
  }
  return duplicateFilenameIds;
}
// Add files with duplicated names to
//  a newly created folder.
function addDuplicateFilesToFolder() {
  var duplicateFilenameIds = 
      getDuplicateFilenameIds(),
      folderName = 'DUPLICATE FILES',
      folder =
       DocsList.createFolder(folderName);
  duplicateFilenameIds.forEach(
    function (fileId) {
      var file = 
        DocsList.getFileById(fileId);
      file.addToFolder(folder);
    });
}

// Chapter 9

// Return an array of cell values from a 
//  named range.
// The call to "map" is required because
//   the array derived from the range is an 
//   array of arrays.
// For this example, the input range is only one
//  column wide so only the first element of each
//  array is needed.
function getEmailList() {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      rngName = 'EmailContacts',
      emailRng = ss.getRangeByName(rngName),
      rngValues = emailRng.getValues(),
      emails = [];
  emails = rngValues.map(
    function (row) {
      return row[0];
    });
  return emails;
}
// Given a Folder object and an array of email
//  addresses, add each email address as an
//  editor for the Folder.
function addEditorsToFolder(folder, editors) {
  editors.forEach(
    function (editor) {
      folder.addEditor(editor);
    });
}
// Create a new folder, add a pre-defined
//  list of editors to the new folder and
//  email each of the editors to inform them.
// Will need to authorise!
function notifyEditors (){
  var folderName = 'Important Files',
      newFolder = 
      DocsList.createFolder(folderName),
      newFolderUrl = newFolder.getUrl(),
      emailList = getEmailList(),
      subject = 'New Folder to Edit',
      body = 
      'A new folder named ' + 
      folderName +
      ' has been created for important files. ' +
      'You can add and edit files in this ' +
        'location.\n' +
        ' The URL is:\n' +
        newFolderUrl;
      addEditorsToFolder(newFolder, 
                         emailList);
      MailApp.sendEmail(emailList.join(','),
                        subject,
                        body);
}
//  To run, replace "<emails go here>" with
//    real email addresses.
//  Creates a document and write some text to it.
//  Saves and closes the document,
//   Beware: forgetting " doc.saveAndClose()" 
//  results in an empty PDF being sent.
//  Use DocsList getFileById() method to 
// return a File object.
// Read this as a PDF using the File getAs()
// method.
// Create the attachment object and send it.
// Requires authorisation.
function sendAttachment() {
  var doc
      = DocumentApp.create('ToSendAsAttachment'),
      emailList = [<emails go here>],
      file,
      pdf,
      pdfName = 'Test.pdf',
      fileId = doc.getId(),
      subject = 'Test attachment',
      body = 'See attached PDF',
      attachment,
      paraText = 
      'This is text that will be written\n' +
      'to a document that will then be saved\n' +
      'as a PDF and sent as an attachment. ';
  doc.appendParagraph(paraText);
  doc.saveAndClose();
  file = DocsList.getFileById(fileId);
  pdf = file.getAs('application/pdf').getBytes();
  attachment = {fileName: pdfName,
                content:pdf, 
                mimeType:'application/pdf'};
  MailApp.sendEmail(emailList.join(','), 
                    subject,
                    body,
                   {attachments:[attachment]});
}
// Return a list of objects that contain
//  selected details of the user's email
//  threads.
// This will be slow for a large mail box
//  so there is an optional argument to limit
//  the returned array to a certain number.
// The most recent threads are returned
//  returned first.
function getThreadSummaries(threadCount) {
  var threads = 
        GmailApp.getInboxThreads(),
      threadSummaries = [],
      threadSummary = {},
      recCount = 0;
  if (typeof threadCount !== 'number') {
    threadCount = GmailApp.getInboxThreads();
  }
  threads.forEach(
    function (thread) {
      recCount += 1;
      if (recCount > threadCount) {
        return;
      }
      threadSummary = {};
      threadSummary['MessageCount'] =
        thread.getMessageCount();
      threadSummary['Subject'] =
        thread.getFirstMessageSubject();
      threadSummary['ThreadId'] =
        thread.getId();
      threadSummary['LastUpdate'] =
        thread.getLastMessageDate();
      threadSummary['URL'] = 
        thread.getPermalink();
      threadSummaries
         .push(threadSummary);
    });
  return threadSummaries;
}
// Write some thread details returned by
//  function "getThreadSummaries()" to
//  a newly inserted sheet named
//  "EmailThreadSummary"
// Make sure the sheet is deleted before
//  re-running or an error will be thrown
//  when it attempts to insert a new sheet
// with the same name.
function writeThreadSummary() {
  var threadCount = 10,
      threadSummaries = 
      getThreadSummaries(threadCount),
      ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.insertSheet(),
      headerRow;
  sh.setName('EmailThreadSummary');
  sh.appendRow(['Subject',
                'MessageCount',
                'LastUpdate',
                'ThreadId',
                'URL']);
  headerRow = sh.getRange(1,1,1,5);
  headerRow.setFontWeight('bold');
  threadSummaries.forEach(
    function (threadSummary) {
      sh.appendRow([threadSummary['Subject'],
                   threadSummary['MessageCount'],
                   threadSummary['LastUpdate'],
                   threadSummary['ThreadId'],
                   threadSummary['URL']]);
    });
}
// Create an array of objects containing
// information extracted from each GmailMessage
// object for the given a thread ID.
// Each object property is populated by
// the return value of a GmailMessage
//  method.
// Return 'undefined' if there are no
// messages.
function getMsgsForThread(threadId) {
  var thread = 
      GmailApp.getThreadById(threadId),
      messages,
      msgSummaries = [],
      messageSummary = {};
  if (!thread) {
    return;
  }
  messages = thread.getMessages();
  messages.forEach(
    function (message) {
      messageSummary = {};
      messageSummary['From'] =
        message.getFrom();
      messageSummary['Cced'] =
        message.getCc();
      messageSummary['Date'] =
        message.getDate();
      messageSummary['Body'] =
        message.getPlainBody();
      messageSummary['MsgId'] =
        message.getId();
      msgSummaries.push(messageSummary);
    });
  return msgSummaries;
}
// Given a message summary object as returned
//  in an array by "getMsgsForThread" and
//  a Google Document ID, open the document,
// extract the object values and write them
// to the document.
// Argument checks could/should be added to
// ensure that the first argument is an
// object and that the second is a valid
// ID for a document.
function writeMsgBodyToDoc(msgSummsForThread,
                           docId) {
  var doc = DocumentApp.openById(docId),
      header,
      from = 
      msgSummsForThread['From'],
      msgDate =
      msgSummsForThread['Date'],
      msgBody = 
      msgSummsForThread['Body'],
      msgId = 
      msgSummsForThread['MsgId'],
      docBody = doc.getBody(),
      paraTitle;
  docBody.appendParagraph('From: ' +
                          from + 
                          '\rDate: ' +
                          msgDate + 
                         '\rMessage ID: ' +
                         msgId);
  docBody.appendParagraph(msgBody);
  docBody
  .appendParagraph('   ####################   ');
  doc.saveAndClose();
}
// Use the Thread IDs generated earlier
//  to extract the message for each thread
// Create a new Google Document
// Write summary message data to the Google
//  Document.
// Contains a nested forEach structure,
//  see explanation in text.
function writeMsgsForThreads() {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      shName = 'EmailThreadSummary',
      sh = ss.getSheetByName(shName),
      docName = 'ThreadMessages',
      doc =
      DocumentApp.create(docName),
      docId = doc.getId(),
      threadIdCount =
      sh.getLastRow() -1,
      rngThreadIds = 
        sh.getRange(2, 
                    4, 
                    threadIdCount,
                    1),
      threadIds = 
        rngThreadIds.getValues();
  threadIds.forEach(
      function(row) {
        var threadId
        = row[0],
        msgsForThread = 
        getMsgsForThread(threadId);
        msgsForThread.forEach(
          function (msg) {
            writeMsgBodyToDoc(
            msg,
            docId);
          });
      });
  
}
// Add a label to all email threads where
//  the first message in the thread has a
// specified text in its subject section.
// To do this:
//  1: Create a label object using the
//     GmailApp createLabel() method.
//  2: Retrieve all the inbox threads.
//  3: Filter the resulting array of
//     threads based on the specified
//     subject text.
//  4: Add the label to each of the
//     threads in the filtered array
//  5: Call the thread refresh method
//     to display the label.
function labelAdd() {
  var subjectText =
      '[Leanpub] A new reader just purchased ' +
        'Google Spreadsheet Programming!',
      labelText = 'BooksBought',
      label = 
      GmailApp.createLabel(labelText),
      threadsAll = 
      GmailApp.getInboxThreads(),
      threadsToLabel;
  threadsToLabel = threadsAll.filter(
    function (thread) {
      return (thread.getFirstMessageSubject()
              === 
              subjectText);
    });
  Logger.log(threadsToLabel.length);
  threadsToLabel.forEach(
    function (thread) {
      thread.addLabel(label);
      thread.refresh();
    });
}
// Using a hard-code thread ID taken from
// the email URL, extract the first message
// from the thread and extract the first
// attachment from that message.
// Copy the attachment into a Blob.
// Use the DocsList object to create a File
// object from this blob.
function putAttachmentInGoogleDrive() {
  var threadId = '140541c39ae7ce97',
      thread = 
      GmailApp.getThreadById(threadId),
      firstMsg = 
      thread.getMessages()[0],
      firstAttachment =
      firstMsg.getAttachments()[0],
      blob = 
      firstAttachment.copyBlob();
  DocsList.createFile(blob);
  Logger.log('File named ' + 
             firstAttachment.getName() +
             ' has been saved to Google Drive');
}
// Write some basic information for all
//  calendars available to the user to
//  the log.
function calendarSummary() {
  var cals = 
      CalendarApp.getAllOwnedCalendars();
  Logger.log('Number of Calendars: '
             + cals.length);
  cals.forEach(
    function(cal) {
      Logger.log('Calendar Name: ' + 
                 cal.getName());
      Logger.log('Is primary calendar? ' +
                 cal.isMyPrimaryCalendar());
      Logger.log('Calendar description: ' +
                 cal.getDescription());
                 
    });
}
// Take a range of dates from 8th August to
//  16th August 2013 (inclusive) from a 
//  spreadsheet and create calendar 
//  all-day events for each of these dates.
// The title is set as "Holidays" and the
//  description to "Forget about work".
function calAddEvents() {
  var cal = CalendarApp.getDefaultCalendar(),
      ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName('Holidays'),
      holDates = 
      sh.getRange('A1:A9').getValues().map(
        function (row) {
          return row[0];
        });
  holDates.forEach(
    function (holDate) {
      var calEvent;
        calEvent  = 
        cal.createAllDayEvent('Holiday',
                              holDate);
        calEvent.setDescription(
                     'Forget about work')
    });
}
// Remove all events from the default calendar
//  for the dates between Aug 8th and Aug 16th
//  inclusive where the event title equals
//  "Holiday".
// First get all events for the date range
//  as an array.
// Then filter this array to
//  get only those with the indicated title.
// Remove those filtered events from the
//  calendar.
function calRemoveEvents() {
  var cal = CalendarApp.getDefaultCalendar(),
      calEvents,
      eventTitleToCancel = 'Holiday',
      toCancelEvents = [];
  calEvents = 
    cal.getEvents(new Date('August 08, 2013'),
                  new Date('August 16, 2013'));
  toCancelEvents = calEvents.filter(
    function (calEvent) {
      return (calEvent.getTitle() 
              ===
             eventTitleToCancel);
    });
  toCancelEvents.forEach(
    function (eventToCancel) {
      eventToCancel.deleteEvent();
    });
}
// Get the default calendar object.
// Insert a new sheet into the active
//  spreadsheet.
// Retrieve all CalendarEvent objects
//  for a specified date interval
//  (month of August 2013).
// Append a header row to the new sheet
//  and populate it.
// Make the header row font bold.
// Loop over the array of CalendarEvent objects
//  and use their methods to extract property
//  values of interest.
// Write the property values to the new sheet.
function writeCalEventsToSheet() {
  var cal = CalendarApp.getDefaultCalendar(),
      ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      newShName = 'CalendarEvents',
      newSh = ss.insertSheet(newShName),
      startDate = new Date('August 01, 2013'),
      endDate = new Date('August 31, 2013'),
      calEvents = cal.getEvents(startDate, endDate),
      colHeaders = [],
      colCount,
      colHeaderRng;
  colHeaders = ['Title',
               'Description',
               'EventId',
               'DateCreated',
               'IsAllDay',
               'MyStatus',
               'StartTime',
               'EndTime',
               'Location'];
  colCount = colHeaders.length;
  newSh.appendRow(colHeaders);
  colHeaderRng = newSh.getRange(1, 
                 1, 
                 1, 
                 colCount);
  colHeaderRng.setFontWeight('bold');
  calEvents.forEach(
    function (calEvent) {
      newSh.appendRow(
       [calEvent.getTitle(),
       calEvent.getDescription(),
       calEvent.getId(),
       calEvent.getDateCreated(),
       calEvent.isAllDayEvent(),
       calEvent.getMyStatus(),
       calEvent.getStartTime(),
       calEvent.getEndTime(),
       calEvent.getLocation()
       ]);
    });
}

// Chapter 10
// For this chapter, HTML and client JavaScript will be enclosed in /* */ comment blocks!

// Execute this function and switch to 
//  spreadsheet tab to see the HTML form
//   built in the file index.html.
function demoHtmlServices() {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      html = 
      HtmlService
        .createHtmlOutputFromFile('index');
  ss.show(html);
}
// This function is called by the JavaScript
//  function "formSubmit()" defined in the
//  accompanying HTML file.
function getValuesFromForm(form){
  var firstName = form.firstName,
      lastName = form.lastName,
      sheet = 
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getActiveSheet();
  sheet.appendRow([firstName, lastName]);
}
/*
<!--
A very simple data entry form that writes
the text input values back to a spreadsheet.
-->
<div>
<b>Add Row To Spreadsheet</b><br />
<form>
First name: <input id="firstname" 
             name="firstName" type="text" />
<br>
Last name: <input id="lastname" 
            name="lastName" type="text" />
<br>
<input onclick="formSubmit()" 
       type="button" value="Add Row" />
<input onclick="google.script.host.close()" 
       type="button" value="Exit" />
</form>
<script type="text/javascript">
function formSubmit() {
  google.script.run.
    getValuesFromForm(document.forms[0]);
        }
    </script>
</div>
*/
/*
<!-- 
Create a simple user feedback form called "survey.html".
-->
<div>
<h1>Customer Satisfaction</h1>
  <form>
    <fieldset>
    <legend>Enter Customer Details:</legend>
    <p><label>Email: </label>
      <input type="text" name="email" size="30"/>
    </p>
	<p><label>Gender: </label>
         <input type="radio" name="gender"
                value="Male" id="gender"/> Male
         <input type="radio" name="gender"
                value="Female" id="gender"/> Female
    </p>
	<p><label>Country:
  <select name="country" id="country">
  <option value="USA">USA</option>
  <option value="Canada">Canada</option>
  <option value="UK">UK</option>
  <option value="Australia">Australia</option>
  <option value="New Zealand">New Zealand</option>
	</select></label>
    </p>
    </fieldset>
    <fieldset>
    <legend class="mylbl">Lengthy Note</legend>
       <textarea rows="4" cols="58" 
         name="note" id="note">

	</textarea> 
   </fieldset>
   <p>
   <input type="button" 
      value="Send Feedback" onclick="formSubmit()"/>
   <input type="button" 
     value="Cancel" onclick="clear()" />
   </p>
   <p id="message">
   </p>
  </form>
<script type="text/javascript">
function formSubmit() {
  google.script.run.
    getValuesFromForm(document.forms[0]);
  document.forms[0].reset();
  alert('Submitted');
}
function clear() {
  document.forms[0].reset();
}
</script>
</div>
*/
// Required function name for web apps.
function doGet() {
  var html =
    HtmlService
     .createHtmlOutputFromFile('survey');
  html.setTitle('Customer Survey');
  html.setHeight(600);
  return html;
}
// Extract values from the web app form
//  and write them to a sheet named
//  "Feedback".
function getValuesFromForm(form){
  var email = form.email,
      gender = form.gender,
      country = form.country,
      note = form.note,
      ssId = 
'0Amdsdq7IKB9ydExEanFqMm5ocmRSMndyRmFOeTgxckE',
      ss = SpreadsheetApp.openById(ssId),
      shName = 'Feedback',
      sheet = ss.getSheetByName(shName);
  sheet.appendRow([email,
                   gender,
                  country,
                  note]);
}
/*
<div>
<h1>English Premier League</h1>

<? var data = getData(); ?>
<table style="border: 1px solid black;">
  <? for (var i = 0; i < data.length; i++) { ?>
    <tr>
      <? for (var j = 0; 
                  j < data[i].length; j++) { ?>
        <td style="border: 1px solid black;">
        <?= data[i][j] ?></td>
      <? } ?>
    </tr>
  <? } ?>
</table>
</div>
*/
function doGet() {
  var html =
  HtmlService
    .createTemplateFromFile('premier_league');
  return html.evaluate();
}
function getData() {
  var ssId = 
  '0Amdsdq7IKB9ydE5KYXkyRmJZZ244Qmo0ODVrX0dXekE',
      ss = SpreadsheetApp.openById(ssId),
      rng = 
    ss.getRangeByName('premier_league_table'),
      data = rng.getValues();
  return data;
}

// Appendix A
// Equivalent VBA code is in /* */ comment blocks

/*
Public Sub SpreadsheetInstance(()
    Dim ss As Workbook
    Set ss = Application.ActiveWorkbook
    Debug.Print ss.Name
End Sub
*/
function spreadsheetInstance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(ss.getName());
}
/*
Public Sub FirstSheetInfo()
    Dim sh1 As Worksheet
    Set sh1 = ActiveWorkbook.Worksheets(1)
    Dim usedRng As Range
    Set usedRng = sh1.UsedRange
    Debug.Print sh1.Name
    Debug.Print usedRng.Address
End Sub
*/
function firstSheetInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets(),
      // getSheets() returns an array
      // JavaScript arrays are always zero-based
      sh1 = sheets[0];
  Logger.log(sh1.getName());
  // getDataRange is analagous to UsedRange 
  //in VBA
  // getA1Notation() is functional equivalent to
  //  Address in VBA
  Logger.log(sh1.getDataRange().getA1Notation());
}
/*
Public Sub PrintSheetNames()
    Dim sheets As Worksheets
    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Sheets
        Debug.Print sheet.Name
    Next sheet
End Sub
*/
// Print the names of all sheets in the active
//  spreadsheet.
function printSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets(),
      i;
  for (i = 0; i < sheets.length; i += 1) {
    Logger.log(sheets[i].getName());
  }
}
/*
' Add a new sheet to a workbook.
' Call the Add method of the 
'  Worksheets collection
' Assign a name to the returned 
'  Worksheet instance
'   Name property.
Sub AddNewSheet()
    Dim newSheet As Worksheet
    Set newSheet = ActiveWorkbook.Worksheets.Add
    newSheet.Name = "AddedSheet"
    MsgBox "New Sheet Added!"
End Sub

' Delete a named sheet from the 
'  active spreadsheet.
' The sheet to delete is identified
'   in the Worksheets collection
' by name. The returned instance 
'  is deleted by calling its
' Delete method.
' MS Excel will prompt to confirm.
Sub RemoveSheet()
    Dim sheetToRemove As Worksheet
    Set sheetToRemove = _
     ActiveWorkbook.Worksheets("AddedSheet")
    sheetToRemove.Delete
    MsgBox "Sheet Deleted!"
End Sub
*/
// Add a new sheet to the active spreadsheet.
// Get an instance of the active spreadsheet.
// Call its insertSheet method.
// Call the setName method of the 
//  returned instance.
function addNewSheet() {
  var ss = 
     SpreadsheetApp.getActiveSpreadsheet(),     
      newSheet;
  newSheet = ss.insertSheet();
  newSheet.setName("AddedSheet");
  Browser.msgBox("New Sheet Added!");
}

// Remove a named sheet from the 
//  active spreadsheet.
// Get an instance of the active 
//  spreadsheet.
// Get an instance of the sheet to remove.
// Activate the sheet to remove
// Call the spreadsheet instance method
//   deleteActiveSheet.
function removeSheet() {
  var ss =
   SpreadsheetApp.getActiveSpreadsheet(),
   sheetToRemove = 
       ss.getSheetByName("AddedSheet");
  sheetToRemove.activate();
  ss.deleteActiveSheet();
  Browser.msgBox("SheetDeleted!");
}
/*
Public Sub SheetHide()
    Dim sh As Worksheet
    Set sh = Worksheets.Item("ToHide")
    sh.Visible = False
End Sub
*/
// Hide a sheet specified by its name.
function sheetHide() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName('ToHide');
  sh.hideSheet()
}
/*
Public Sub ListHiddenSheetNames()
    Dim sheet As Worksheet
    For Each sheet In Worksheets
        If sheet.Visible = False Then
            Debug.Print sheet.Name
        End If
    Next sheet
End Sub
*/
// Write a list of hidden sheet names to log.
function listHiddenSheetNames() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets();
  sheets.forEach(
    function (sheet) {
      if (sheet.isSheetHidden()) {
        Logger.log(sheet.getName());
      }
    });
}
/*
Public Sub SheetsUnhide()
    Dim sheet As Worksheet
    For Each sheet In Worksheets
        If sheet.Visible = False Then
            sheet.Visible = True
        End If
    Next sheet
End Sub
*/
// Unhide all hidden sheets.
function sheetsUnhide() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets();
  sheets.forEach(
    function (sheet) {
      if (sheet.isSheetHidden()) {
        sheet.showSheet();
      }
    });
}
/*
' Password-protect rotect a sheet identified
' by name
Public Sub SheetProtect()
    Dim sh As Worksheet
    Dim pwd As String: pwd = "secret"
    Set sh = Worksheets.Item("ToProtect")
    sh.Protect pwd
End Sub
*/
// Identify a sheet by name to protect
// When this code runs, the lock icon
// will appear on the sheet name.
// Share the spreadsheet with another user
// as an editor. That user can edit all
// sheets except the protected one. The user
// can still edit the protected sheet.
function sheetProtect() {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName('ToProtect'),
      permissions = sh.getSheetProtection();
  ss.addEditor(<gmail address goes here>);
  permissions.setProtected(true);
  sh.setSheetProtection(permissions);
}
/*
Public Sub PrintSelectionDetails()
    Debug.Print "Selected Range Details: "
    Debug.Print "-- Sheet: " & _
                Selection.Worksheet.Name
    Debug.Print "-- Address: " & _
 
      Selection.Address
    Debug.Print "-- Row Count: " & _
                Selection.Rows.Count
    Debug.Print "'-- Column Count: " & _
                 Selection.Columns.Count
End Sub
*/
// Prints details about selected range in 
//  active spreadsheet
// To run, paste code into script editor,
//   select some cells on any sheet, 
//    execute code and
//   check log to see details
// Prints details about selected range 
//  in active spreadsheet
// To run, paste code into script editor,
//   select some cells on any sheet, 
//  execute code and
//   check log to see details
function printSelectionDetails() {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      selectedRng = ss.getActiveRange();
  Logger.log('Selected Range Details:');
  Logger.log('-- Sheet: '
             + selectedRng
                .getSheet()
                .getSheetName());
  Logger.log('-- Address: '
             + selectedRng.getA1Notation());
  Logger.log('-- Row Count: ' 
             + ((selectedRng.getLastRow() + 1) 
             - selectedRng.getRow()));
  Logger.log('-- Column Count: ' 
             + ((selectedRng.getLastColumn() + 1)
             - selectedRng.getColumn()));
}
/*
Public Function GetUsedRangeAsArray(sheetName  _
                          As String) As Variant
    Dim sh As Worksheet
    Set sh = _
      ActiveWorkbook.Worksheets(sheetName)
    GetUsedRangeAsArray = sh.UsedRange.value
End Function
Sub test_GetUsedRangeAsArray()
    Dim sheetName As String
    Dim rngValues
    Dim firstRow As Variant
    sheetName = "Sheet1"
    rngValues = GetUsedRangeAsArray(sheetName)
    Debug.Print rngValues(1, 1)
    Debug.Print UBound(rngValues)
    Debug.Print UBound(rngValues, 2)
End Sub
*/
function getUsedRangeAsArray(sheetName) {
  var ss = 
    SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName(sheetName);
  // The getValues() method of the
  //   Range object returns an array of arrays
  return sh.getDataRange().getValues();
}
// JavaScript does not distinguish between
//  subroutines and functions.
// When the return statement is omitted,
//  functions return undefined.
function test_getUsedRangeAsArray() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheetName = 'Sheet1',
      rngValues = getUsedRangeAsArray(sheetName);
  // Print the number of rows in the range
  // The toString() call to suppress the 
  // decimal point so
  //  that, for example, 10.0, is reported as 10
  Logger.log((rngValues.length).toString());
  // Print the number of columns
  // The column count will be the same 
  // for all rows so only need the first row
  Logger.log((rngValues[0].length).toString());
  // Print the value in the first cell
  Logger.log(rngValues[0][0]);
}
/*
Sub AddColorsToRange()
    Dim sh1 As Worksheet
    Dim addr As String: addr = "A4:B10"
    Set sh1 = ActiveWorkbook.Worksheets(1)
    sh1.Range(addr).Interior.ColorIndex = 3
    sh1.Range(addr).Font.ColorIndex = 10
End Sub
*/
// Select a block of cells in the first sheet.
// Use Range methods to set both the font and
//   background colors.
function addColorsToRange() {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets(),
      sh1 = sheets[0],
      addr = 'A4:B10',
      rng;
  // getRange is overloaded. This method can
  //  also accept row and column integers
  rng = sh1.getRange(addr);
  rng.setFontColor('green');
  rng.setBackgroundColor('red');
}
/*
Public Sub OffsetDemo()
    Dim sh As Worksheet
    Dim cell As Range
    Set sh = _
      ActiveWorkbook.Worksheets(1)
    Set cell = sh.Range("B2")
    cell.value = "Middle"
    cell.Offset(-1, -1).value = "Top Left"
    cell.Offset(0, -1).value = "Left"
    cell.Offset(1, -1).value = "Bottom Left"
    cell.Offset(-1, 0).value = "Top"
    cell.Offset(1, 0).value = "Bottom"
    cell.Offset(-1, 1).value = "Top Right"
    cell.Offset(0, 1).value = "Right"
    cell.Offset(1, 1).value = "Bottom Right"
End Sub
*/
// The Spreadsheet method getSheets() returns
//  an array.
// The code "ss.getSheets()[0]"
//  returns the first sheet and is equivalent to
// "ActiveWorkbook.Worksheets(1)" in VBA.
// Note that the VBA version is 1-based!
function offsetDemo() {
  var ss =
   SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheets()[0],
      cell = sh.getRange('B2');
  cell.setValue('Middle');
  cell.offset(-1,-1).setValue('Top Left');
  cell.offset(0, -1).setValue('Left');
  cell.offset(1, -1).setValue('Bottom Left');
  cell.offset(-1, 0).setValue('Top');
  cell.offset(1, 0).setValue('Bottom');
  cell.offset(-1, 1).setValue('Top Right');
  cell.offset(0, 1).setValue('Right');
  cell.offset(1, 1).setValue('Bottom Right');
}
/*
' Mimicking Google Apps Script
'  offset() method overloads.
Public Sub OffsetOverloadDemo()
    Dim sh As Worksheet
    Dim cell As Range
    Dim offsetRng2 As Range
    Dim offsetRng3 As Range
    Set sh = ActiveWorkbook.Worksheets(1)
    Set cell = sh.Range("A1")
    'Offset returns a Range so Offset 
    ' can be called again
    ' on the returned Range from 
    '  first Offset call.
    Set offsetRng2 = Range(cell.Offset(1, 4), _
                   cell.Offset(1, 4).Offset(1, 0))
    Set offsetRng3 = Range(cell.Offset(10, 4), _
                   cell.Offset(10, 4).Offset(3, 4))
    Debug.Print offsetRng2.Address
    Debug.Print offsetRng3.Address
End Sub
*/
// Demonstrating overloaded versions of offset()
// Output:
// Address of offset() overload 2 
//  (rowOffset, columnOffset, numRows) is: E2:E3
//  Address of offset() overload 3 (rowOffset, 
//    columnOffset, numRows, numColumns)
//     is: E11:I14
function offsetOverloadDemo() {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheets()[0],
      cell = sh.getRange('A1'),
      offsetRng2 = cell.offset(1, 4, 2),
      offsetRng3 = cell.offset(10, 4, 4, 5);
  Logger.log('Address of offset() overload 2 ' +
         '(rowOffset, columnOffset, numRows) is: ' 
             + offsetRng2.getA1Notation());
  Logger.log('Address of offset() overload 3 ' +
         '(rowOffset, columnOffset, numRows, ' +
         'numColumns) is: '
         + offsetRng3.getA1Notation());
}
/*
Public Sub PrintRangeNames()
  Dim namedRng As Name
  For Each namedRng In ActiveWorkbook.Names
    Debug.Print "The name of the range is: " & _
         namedRng.Name & _
        " It refers to this address: " & _
           
         namedRng.RefersTo
    Next namedRng
End Sub
*/
/*
Public Sub SetCellComment(sheetName As String, _
                        cellAddress As String, _ 
                        
                        cellComment As String)
    Dim sh As Worksheet
    Dim cell As Range
    Set sh = ActiveWorkbook.Worksheets(sheetName)
    Set cell = sh.Range(cellAddress)
    cell.AddComment cellComment
End Sub
Public Sub test_SetCellComment()
    Dim sheetName As String
    sheetName = "Sheet1"
    Dim cellAddress As String
    cellAddress = "C10"
    Dim cellComment As String
    cellComment = "Comment added: " & Now()
    Call SetCellComment(sheetName, _
                        
    cellAddress, _
                        
    cellComment)
End Sub
*/
function setCellComment(sheetName, cellAddress,
                        cellComment) {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName(sheetName),
      cell = sh.getRange(cellAddress);
  cell.setNote(cellComment);
}
function test_setCellComment() {
  var sheetName = 'Sheet1',
      cellAddress = 'C10',
      cellComment = 'Comment added ' + Date();
  setCellComment(sheetName, cellAddress, cellComment);
}
/*
' Need to check if the cell has a comment.
' If it does not, then exit the sub but if
' it does, then remove it.
Public Sub RemoveCellComment(sheetName _
                             As String, _
                    cellAddress As String)
    Dim sh As Worksheet
    Dim cell As Range
    Set sh = ActiveWorkbook.Worksheets(sheetName)
    Set cell = sh.Range(cellAddress)
    If cell.Comment Is Nothing Then
        Exit Sub
    Else
        cell.Comment.Delete
    End If
End Sub
Public Sub test_RemoveCellComment()
    Dim sheetName As String
    sheetName = "Sheet1"
    Dim cellAddress As String
    cellAddress = "C10"
    Call RemoveCellComment(sheetName, _
                           cellAddress)
End Sub
*/
// To remove a comment, just pass an empty string
//  to the setNote() method.
function removeCellComment(sheetName, cellAddress) {
  var ss =
      SpreadsheetApp.getActiveSpreadsheet(),
      sh = ss.getSheetByName(sheetName),
      cell = sh.getRange(cellAddress);
  cell.setNote('');
}
function test_removeCellComment() {
  var sheetName = 'Sheet1',
      cellAddress = 'C10';
  removeCellComment(sheetName, cellAddress);
}
/*
' This VBA code is commented because the 
'  VBA approach differs
'   considerably from the Google Apps Script one.
' Note: the Offset() method of the Range 
'  object uses 0-based indexes.
Public Sub copyRowsToNewSheet()
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    newSheetName = "Target"
    Dim sourceRng As Range
    Dim sourceRows As Variant
    Dim i As Long
    Set sourceSheet = _ 
        Application.Worksheets("Source")
    Set newSheet = ActiveWorkbook.Worksheets.Add
    newSheet.Name = newSheetName
    ' Use a named range as marker 
    '  for row copying (VBA hack!)
    newSheet.Range("A1").Name = "nextRow"
    Set sourceRng = sourceSheet.UsedRange
    ' Copy the header row
    sourceRng.Rows(1).Copy Range("nextRow")
    ' Moved the named range marker down one row
    Range("nextRow").Offset(1, 0).Name = _
       "nextRow"
    'Skip header row by setting i, 
    ' the row counter, = 2
    ' i starts at 2 to skip header row
    For i = 2 To sourceRng.Rows.Count
        If sourceRng.Cells(i, 2).value _
            <= 10000 Then
            ' Define the row range to copy 
            ' using the first and
            '   last cell in the row.
            Range(sourceRng.Cells(i, 1), _ 
                  sourceRng.Cells(i, _
            sourceRng.Columns.Count)).Copy _
              Range("nextRow")
            Range("nextRow").Offset(1, 0).Name _
             = "nextRow"
        End If
    Next i
End Sub
*/
// Longer example
// Copy rows from one sheet named "Source" to
//  a newly inserted
//   one based on a criterion check of second
//   column.
// Copy the header row to the new sheet.
// If Salary <= 10,000 then copy the entire row
function copyRowsToNewSheet() {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sourceSheet = ss.getSheetByName('Source'),
      newSheetName = 'Target',
      newSheet = ss.insertSheet(newSheetName),
      sourceRng = sourceSheet.getDataRange(),
      sourceRows = sourceRng.getValues(),
      i;
  newSheet.appendRow(sourceRows[0]);
  for (i = 1; i < sourceRows.length; i += 1) {
    if (sourceRows[i][1] <= 10000) {
      newSheet.appendRow(sourceRows[i]);
    }
  } 
}
/*
Public Sub test_PrintSheetFormulas()
    Dim sheetName As String
    sheetName = "Formulas"
    Call PrintSheetFormulas(sheetName)
End Sub
Public Sub PrintSheetFormulas(sheetName _
                               
  As String)
    Dim sourceSheet As Worksheet
    Dim usedRng As Range
    Dim i As Long
    Dim j As Long
    Dim cellAddr As String
    Dim cellFormula As String
    Set sourceSheet = _
       
    ActiveWorkbook.Worksheets(sheetName)
    Set usedRng = sourceSheet.UsedRange
    For i = 1 To usedRng.Rows.Count
        For j = 1 To usedRng.Columns.Count
            cellAddr = _
              
            usedRng.Cells(i, j).Address
            cellFormula = _
              
            usedRng.Cells(i, j).Formula
            If Left(cellFormula, 1) = "=" Then
                Debug.Print cellAddr & _
                  
                ": " & cellFormula
            End If
        Next j
    Next i
End Sub
*/
function test_printSheetFormulas() {
  var sheetName = 'Formulas';
  printSheetFormulas(sheetName);
}
function printSheetFormulas(sheetName) {
  var ss = 
      SpreadsheetApp.getActiveSpreadsheet(),
      sourceSheet = ss.getSheetByName(sheetName),
      usedRng = sourceSheet.getDataRange(),
      i,
      j,
      cell,
      cellAddr,
      cellFormula;
  for (i = 1; i <= usedRng.getLastRow();
            i += 1) {
    for (j = 1; j <= usedRng.getLastColumn(); 
                 j += 1) {
      cell = usedRng.getCell(i, j);
      cellAddr = cell.getA1Notation();
      cellFormula = cell.getFormula();
      if (cellFormula) {
        Logger.log(cellAddr + 
          ': ' + cellFormula);
      }
    }
  }
}


 
