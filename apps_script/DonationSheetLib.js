//
// Copyright 2025 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This library includes functions to handle import of files containing donations made to the Newton Tree
// Conservancy. These functions exist primarily to normalize incoming data to the layout captured in the
// spreadsheets used to manage donations and membership.
//
// @OnlyCurrentDoc
//
const DEPLOYMENT_ID                           = "1cXoHvwTUh5pTV3_0YHl9jZsL4YZ7Ie6juG307YwOBxGLjeF81khFYHcy";
const DEPLOYMENT_VERSION                      = "1";
const NEWTON_TREE_CONSERVANCY_MENU            = "Newton Tree Conservancy";
const ABOUT_MENU_ITEM                         = "About...";
const PROCESS_PENDING_DONATION_DATA_MENU_ITEM = "Process pending donation data";
const PROCESS_PENDING_DONATION_DATA_TITLE     = "Archive Data for Planting Date";
const ABOUT_TITLE                             = "About Donation Processing Spreadsheet";

const TEST_CVS_URL = "https://drive.google.com/file/d/1Qbl5It6wE9BSiSMqXE4lIYL_qUC9szGV/view?usp=drive_link";
const TEST_CVS_ID  = "1Qbl5It6wE9BSiSMqXE4lIYL_qUC9szGV";

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();

  ui
    .createMenu(NEWTON_TREE_CONSERVANCY_MENU)
      .addItem(PROCESS_PENDING_DONATION_DATA_MENU_ITEM, "onArchiveDataForPlantingDate")
      .addSeparator()
      .addItem(ABOUT_MENU_ITEM, "onAboutThis")
      .addToUi();
}

function onProcessPendingDonationData(fileId = TEST_CVS_ID){
  let csvFile = DriveApp.getFileById(fileId);

  digestPayPalCsv_(csvFile);
}

function onAboutThis() {
  let ui = SpreadsheetApp.getUi();

  ui.alert(ABOUT_TITLE,
    "Deployment ID\n " + DEPLOYMENT_ID + "\n\n" +
    "Version\n" + DEPLOYMENT_VERSION + "\n\n\n" +
    "Newton Tree Conservancy\n" +
    "www.newtontreeconservancy.org",
    ui.ButtonSet.OK);
}

function digestPayPalCsv_(csvFile) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

  // remove header
  data.splice(0, 1);

  // get coordinates for next available range sheet 
  let startRow = 1;
  let startCol = 1;

  // transform incoming data to correct footprint
  data = transformData_(data);

  // determine size of incoming data
  let numRows = data.length;
  let numColumns = data[0].length;

  // appends data to the sheet.
  sheet.getRange(startRow, startCol, numRows, numColumns).setValues(data);
}

function transformData_(data) {
  return data;
}
