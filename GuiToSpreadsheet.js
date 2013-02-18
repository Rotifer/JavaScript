/*global SpreadsheetApp: false, UiApp: false */

// Close the current UI
function exit() {
    "use strict";
    var ui = UiApp.getActiveApplication();
    return ui.close();
}

// Callback action for button labelled "Add Row"
// Take the values from two text boxes in the active GUI and
//   add them as rows to the active sheet.
// Prepare for next input by:
//   re-setting the text boxes to empty strings
//   Set the focus to the first text box
function addRow(e) {
    "use strict";
    var ui = UiApp.getActiveApplication(),
        sheet = SpreadsheetApp.getActiveSheet(),
        name = e.parameter.txtName_Name,
        email = e.parameter.txtEmail_Name,
        txtName = ui.getElementById('txtName_Id'),
        txtEmail = ui.getElementById('txtEmail_Id');
    sheet.appendRow([name, email]);
    txtName.setValue('');
    txtEmail.setValue('');
    txtName.setFocus(true);
    return ui;
}

// Build a GUI with two labels, two text boxes and two buttons.
function guiDemo() {
    "use strict";
    var ui = UiApp.createApplication(),
        ss = SpreadsheetApp.getActiveSpreadsheet(),
        uiTitle = 'Add Row To Spreadsheet',
        panelInput = ui.createVerticalPanel(),
        panelName = ui.createHorizontalPanel(),
        panelEmail = ui.createHorizontalPanel(),
        panelButtons = ui.createHorizontalPanel(),
        lblName = ui.createLabel('Name:'),
        lblEmail = ui.createLabel('Email:'),
        txtName = ui.createTextBox(),
        txtEmail = ui.createTextBox(),
        btnAddRow = ui.createButton('Add Row'),
        btnExit = ui.createButton('Exit'),
        exitHandler = ui.createServerHandler('exit'),
        addRowHandler = ui.createServerHandler('addRow');
    panelName.add(lblName);
    panelName.add(txtName);
    panelEmail.add(lblEmail);
    panelEmail.add(txtEmail);
    panelInput.add(panelName);
    panelInput.add(panelEmail);
    panelButtons.add(btnAddRow);
    panelButtons.add(btnExit);
    panelInput.add(panelButtons);
    ui.add(panelInput);
    ui.setWidth(200);
    ui.setHeight(100);
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