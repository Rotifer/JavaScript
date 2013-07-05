/*global SpreadsheetApp: false, Browser: false, DocumentApp: false, Logger: false*/

/* 
 * Written by: mick@javascript-spreadsheet-programming.com
 * 
 * Date: 2013-07-05
 * 
 * Passed by JSLint
 */

/**
 * Copy array values to a document table
 * 
 * Given a document name and an array of arrays, creates the document table.
 * Highlights the first row (header) in bold.
 * Checks for two arguments, the first must be a string and the second an array.
 * 
 * @param {type} docName
 * @param {type} values
 * @returns {undefined}
 */
function writeTableToDocument(docName, values) {
    "use strict";
    var doc,
        table,
        headerRow,
        styles = {};
    if (typeof docName !== 'string') {
        throw {name: 'TypeError',
            message: 'Function writeTableToDocument() ' +
                    'expects a string name for the new ' +
                    'document as its first argument!'};
    }
    if (!Array.isArray(values)) {
        throw {name: 'TypeError',
            message: 'Function writeTableToDocument() ' +
                    'expects an array of values as its second argument.'};
    }
    try {
        doc = DocumentApp.create(docName);
        table = doc.getBody().appendTable(values);
        headerRow = table.getRow(0);
        styles[DocumentApp.Attribute.BOLD] = true;
        headerRow.setAttributes(styles);
        doc.saveAndClose();
    } catch (ex) {
        throw ex;
    }
}

/**
 * Return the cell values of a named range
 * checks that the given argument is type string.
 * Uses this name to reference a spreadsheet range.
 * If the range name does not exists, it will return 'undefined'
 * 
 * @param rngName string
 * @returns array
 */
function getRangeNameValues(rngName) {
    "use strict";
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        namedRng,
        namedRngValues = [];
    if (typeof rngName !== 'string') {
        throw {name: 'TypeError',
            message: 'Function getRangeNameValues() ' +
                    'expects a single string argument for the ' +
                    'name of the target range!'};
    }
    try {
        namedRng = ss.getRangeByName(rngName);
        namedRngValues = namedRng.getValues();
        return namedRngValues;
    } catch (ex) {
        throw ex;
    }
}


/**
 * Runs the copying code
 * 
 * @returns {undefined}
 */
function main() {
    "use strict";
    var rngName = "ContactDetails",
        docName = rngName,
        rngValues;
    try {
        rngValues = getRangeNameValues(rngName);
        writeTableToDocument(docName, rngValues);
        Browser.msgBox('New file created');
    } catch (ex) {
        Browser.msgBox('There has been an error, check the log');
        Logger.log('ERROR:');
        Logger.log(ex.message);
        throw ex;
    }
}

