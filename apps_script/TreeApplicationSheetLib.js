//
// Copyright 2023 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This file includes functions to help manage and normalize data in the spreadsheet that is synchronized
// by a Google Form that residents of Newton use to submit applications for trees to be planted on their
// property. That sheet provides data to the sheets that Newton Tree Conservancy directors use to manage
// planting groups.
//
// @OnlyCurrentDoc
//
const DEPLOYMENT_ID                            = "1eNq3Z-0DFAqclht8OvXxPIM2IvR3J_Q1s4dzaZVERPYyVVB707MVdFPw";
const DEPLOYMENT_VERSION                       = "14";
const HEADER_ROW                               = 2;
const PLANTING_DATE_RANGE                      = "planting_date";
const GROUP_NAME_RANGE                         = "group_name";
const FIRST_NAME_RANGE                         = "first_name";
const LAST_NAME_RANGE                          = "last_name";
const APPL_ACK_EMAIL_SENDER_NAME_RANGE         = "application_ack_email_sender_name"
const APPL_ACK_EMAIL_REPLY_TO_RANGE            = "application_ack_email_reply_to";
const APPL_ACK_EMAIL_SUBJECT_RANGE             = "application_ack_email_subject";
const APPL_ACK_EMAIL_BODY_RANGE                = "application_ack_email_body";
const EMAIL_ADDRESS_RANGE                      = "email_address";
const PLANTER_FIRST_NAME_RANGE                 = "planter_first_name";
const PLANTER_LAST_NAME_RANGE                  = "planter_last_name";
const PLANTER_EMAIL_ADDRESS_RANGE              = "planter_email_address";
const NUMBER_OF_TREES_REQUESTED_RANGE          = "number_of_trees_requested";
const NEWTON_TREE_CONSERVANCY_MENU             = "Newton Tree Conservancy";
const ABOUT_MENU_ITEM                          = "About...";
const ARCHIVE_DATA_FOR_PLANTING_DATE_MENU_ITEM = "Archive data for planting date";
const ARCHIVE_DATA_FOR_PLANTING_DATE_TITLE     = "Archive Data for Planting Date";
const COUNT_OF_ROW_ARCHIVED_TITLE              = "Count of Rows Archived";
const ABOUT_TITLE                              = "About Community Tree Planting Spreadsheet";

const STREET_SUFFIXES = [ 
  ["Ave",   "Avenue"], 
  ["Cir",   "Circle"], 
  ["Ln",    "Lane"], 
  ["Pk",    "Park"], 
  ["Prk",   "Park"], 
  ["Pl",    "Place"], 
  ["Rd",    "Road"], 
  ["Sq",    "Square"],
  ["St",    "Street"],
  ["Ter",   "Terrace"],
  ["Terr",  "Terrace"],
  ["Wy",    "Way"]
];

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();

  ui
    .createMenu(NEWTON_TREE_CONSERVANCY_MENU)
      .addItem(ARCHIVE_DATA_FOR_PLANTING_DATE_MENU_ITEM, "onArchiveDataForPlantingDate")
      .addSeparator()
      .addItem(ABOUT_MENU_ITEM, "onAboutThis")
      .addToUi();
}

function onEdit(e) {
  let sheet = e.source.getActiveSheet();
  let range = sheet.getRange(PLANTING_DATE_RANGE);

  if (sheet.getSheetId() == range.getSheet().getSheetId()) {
    let isLegalValue = true;

    if ((e.value != undefined) && (e.range.rowEnd > 1) && (range.getLastColumn() == e.range.columnEnd)) {
      let parts = e.value.trim().split(/\s+/);

      if (parts.length == 2) {
        if (!Number.isNaN(Number.parseInt(parts[0])) && (parts[0].length == 4)) {
          if ((parts[1] == "Spring") || (parts[1] == "Fall")) {
            e.range.setValue(parts[0] + " " + parts[1]);
          }
          else {
            isLegalValue = false;
          }
        }
        else if (!Number.isNaN(Number.parseInt(parts[1])) && (parts[1].length == 4)) {
          if ((parts[0] == "Spring") || (parts[0] == "Fall")) {
            e.range.setValue(parts[1] + " " + parts[0]);
          }
          else {
            isLegalValue = false;
          }
        }
        else {
          isLegalValue = false;
        }
      }
      else {
        isLegalValue = false;
      }
    }

    if (!isLegalValue) {
      let ui         = SpreadsheetApp.getUi();
      let columnName = sheet.getRange(1, range.getLastColumn()).getValue();

      ui.alert("Invalid Value Specified for " + columnName,
        "Value \"" + e.value + "\" is invalid. Please specify \"YYYY\" followed by \"Spring\" or \"Fall\" with one space between the year and season, and the first letter of the season capitalized.\n\nExample: 2024 Spring",
        ui.ButtonSet.OK);

      e.range.setValue("");    
    }
  }
}

function onSubmit(e) {
  let sheet     = e.range.getSheet();
  let rowIndex  = e.range.getRow();
  let cellRange = sheet.getRange(rowIndex, sheet.getRange(GROUP_NAME_RANGE).getColumn());
  let cellValue = cellRange.getValue();

  cellValue = cellValue.toLowerCase().
    replaceAll("group", "").
    trim();

  let cellParts = cellValue.split(/\s+/);
  let cellIndex = 0;
  let cellToken = "";

  cellValue = "";

  cellParts.forEach(function(e) {
    if (cellIndex > 0) {
      cellValue += " ";  
    }

    cellToken = e.charAt(0).toUpperCase() + e.slice(1);

    cellIndex++;

    if (cellIndex == cellParts.length) {
      cellToken = cellToken.replaceAll(".", "");

      let cellSuffix = STREET_SUFFIXES.find((s) => (s[0] == cellToken));

      if (cellSuffix != undefined) {
        cellToken = cellSuffix[1];
      }
    }

    cellValue += cellToken;
  });

  cellRange.setValue(cellValue);

  cellRange = sheet.getRange(rowIndex, sheet.getRange(NUMBER_OF_TREES_REQUESTED_RANGE).getColumn());
  cellValue = cellRange.getValue();

  if (cellValue == "") {
    cellRange.setValue(0);
  }

  let applicantContactDataRanges = [
    sheet.getRange(rowIndex, sheet.getRange(FIRST_NAME_RANGE).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(LAST_NAME_RANGE).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(EMAIL_ADDRESS_RANGE).getColumn())
  ];
  let planterContactDataRanges = [
    sheet.getRange(rowIndex, sheet.getRange(PLANTER_FIRST_NAME_RANGE).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(PLANTER_LAST_NAME_RANGE).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(PLANTER_EMAIL_ADDRESS_RANGE).getColumn())
  ];

  if (applicantContactDataRanges.every((e, i) => e.getValue().toLowerCase().trim() == planterContactDataRanges[i].getValue().toLowerCase().trim())) {
    planterContactDataRanges.forEach((r) => r.setValue(""));
  }

  sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setVerticalAlignment("top");

  let emailAddress = sheet.getRange(rowIndex, sheet.getRange(EMAIL_ADDRESS_RANGE).getColumn()).getValue();

  if (emailAddress != undefined) {
    let hits = sheet.getDataRange().createTextFinder(emailAddress).matchEntireCell(true).findAll();

    if (hits.length > 1) {
      hits.forEach(function(h) {
        sheet.getRange(h.getRow(), 1, 1, sheet.getLastColumn()).setFontWeight("bold");
      });
    }

    let senderName = sheet.getRange(APPL_ACK_EMAIL_SENDER_NAME_RANGE).getValue();
    let replyTo    = sheet.getRange(APPL_ACK_EMAIL_REPLY_TO_RANGE).getValue();
    let subject    = sheet.getRange(APPL_ACK_EMAIL_SUBJECT_RANGE).getValue();
    let body       = sheet.getRange(APPL_ACK_EMAIL_BODY_RANGE).getValue();

    MailApp.sendEmail(
      emailAddress,
      subject,
      null,
      {
        htmlBody: body,
        replyTo : replyTo,
        name    : senderName
      }
    );
  }
}

function onArchiveDataForPlantingDate() {
  let ui = SpreadsheetApp.getUi();

  let response = ui.prompt(ARCHIVE_DATA_FOR_PLANTING_DATE_TITLE,
    "Enter the planting date you want to archive. Please specify \"YYYY\" followed by \"Spring\" or \"Fall\" with one space between the year and season, and the first letter of the season capitalized.\n\nExample: 2024 Spring",
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let plantingDate = response.getResponseText();

    let file  = SpreadsheetApp.getActiveSpreadsheet();
    let range = file.getRange(PLANTING_DATE_RANGE);
    let sheet = range.getSheet();
    let hits  = range.createTextFinder(plantingDate).matchEntireCell(true).findAll();

    if (hits.length > 0) {
      let deletions = [];
      let archive   = file.insertSheet(plantingDate + " " + "Archive", file.getSheets().length);

      sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).copyTo(archive.getRange("A1"));

      let dstRow = 2;

      hits.forEach(function(h) {
        let srcRow = h.getRow();

        sheet.getRange(srcRow, 1, 1, sheet.getLastColumn()).copyTo(archive.getRange(dstRow, 1));
        deletions.push(srcRow);
        dstRow++;
      });

      deletions.reverse().forEach(d => sheet.deleteRow(d));
    }

    ui.alert(COUNT_OF_ROW_ARCHIVED_TITLE,
      "Number of rows archived is " + hits.length + ".",
      ui.ButtonSet.OK);
  }
}

function onAboutThis() {
  let ui = SpreadsheetApp.getUi();

  ui.alert(ABOUT_TITLE,
    "Deployment ID\n" + DEPLOYMENT_ID + "\n\n" +
    "Version\n" + DEPLOYMENT_VERSION + "\n\n\n" +
    "Newton Tree Conservancy\n" +
    "www.newtontreeconservancy.org",
    ui.ButtonSet.OK);
}
