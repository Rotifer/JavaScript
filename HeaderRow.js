'use strict';
/*global SpreadsheetApp: false */
function HeaderRow(spreadsheet, sheetName, headerRowNumber, startColumnNumber, columnTitles, overwritePrevious) {
  if (arguments.length !== 6) {
    throw {'name': 'Error',
           'message': '"HeaderRow()" constructor function requires 6 arguments!'};
  }
  this.spreadsheet = spreadsheet;
  this.sheetName = sheetName;
  this.headerRowNumber = headerRowNumber;
  this.startColumnNumber = startColumnNumber;
  this.columnTitles = columnTitles;
  this.overwritePrevious = overwritePrevious;
  this.sheet = this.spreadsheet.getSheetByName(this.sheetName);
  this.columnTitleCount = this.columnTitles.length;
  this.headerRowRange = this.sheet.getRange(this.headerRowNumber,
                                            this.startColumnNumber,
                                            1,
                                            this.columnTitleCount);
  this.headerRowRange.setFontWeight('normal');
  this.headerRowRange.setFontStyle('normal');
  this.addColumnTitlesToHeaderRow();
}

HeaderRow.prototype = {
  constructor: 'HeaderRow',
  freezeHeaderRow: function () {
    var sheet = this.sheet;
    sheet.setFrozenRows(this.headerRowNumber);
  },
  setHeaderFontWeightBold: function () {
    this.headerRowRange.setFontWeight('bold');
  },
  setFontStyle: function (style) {
    this.headerRowRange.setFontStyle(style);
  },
  addCommentToColumn: function (comment, headerRowColumnNumber) {
    var cellToComment = this.headerRowRange.getCell(1, headerRowColumnNumber);
    cellToComment.setNote(comment);
  },
  addColumnTitlesToHeaderRow: function () {
    var i,
      titleCell;
    this.spreadsheet.setNamedRange(this.headerRowRangeName, this.headerRowRange);
    for (i = 1; i <= this.columnTitleCount; i += 1) {
      titleCell = this.headerRowRange.getCell(1, i);
      if (titleCell.getValue() && !this.overwritePrevious) {
        throw {'name': 'Error',
               'message': '"HeaderRow.addColumnTitlesToHeaderRow()" Cannot overwrite previous values!'};
      }
      titleCell.setValue(this.columnTitles[i - 1]);
    }
  },
  setHeaderRowName: function (rngName) {
    this.spreadsheet.setNamedRange(rngName, this.headerRowRange);
  }
};

function test_HeaderRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheetName = ss.getActiveSheet().getSheetName(),
    headerRowNumber = 3,
    startColumnNumber = 2,
    columnTitles = ['col1', 'col2', 'col3'],
    overwritePrevious = true,
    hr = new HeaderRow(ss, sheetName, headerRowNumber, startColumnNumber, columnTitles, overwritePrevious);
  hr.freezeHeaderRow();
  hr.setHeaderFontWeightBold();
  hr.setFontStyle('oblique');
  hr.addCommentToColumn('Comment added ' + Date(), 2);
  hr.setHeaderRowName('header');
}
