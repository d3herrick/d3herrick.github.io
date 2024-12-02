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
const deploymentId                       = "1eNq3Z-0DFAqclht8OvXxPIM2IvR3J_Q1s4dzaZVERPYyVVB707MVdFPw";
const deploymentVersion                  = "10";
const plantingDateRange                  = "planting_date";
const groupNameRange                     = "group_name";
const firstNameRange                     = "first_name";
const lastNameRange                      = "last_name";
const applicationAckEmailSenderNameRange = "application_ack_email_sender_name"
const applicationAckEmailReplyToRange    = "application_ack_email_reply_to";
const applicationAckEmailSubjectRange    = "application_ack_email_subject";
const applicationAckEmailBodyRange       = "application_ack_email_body";
const emailAddressRange                  = "email_address";
const planterFirstNameRange              = "planter_first_name";
const planterLastNameRange               = "planter_last_name";
const planterEmailAddressRange           = "planter_email_address";
const numberOfTreeRequestedRange         = "number_of_trees_requested";
const newtonTreeConservancyMenu          = "Newton Tree Conservancy";
const aboutThisMenuItem                  = "About...";
const archiveDataForPlantingDateMenuItem = "Archive data for planting date";
const archiveDataForPlantingDateTitle    = "Archive Data for Planting Date";
const countOfRowsArchivedTitle           = "Count of Rows Archived";
const aboutTitle                         = "About Community Tree Planting Spreadsheet";

const streetSuffixes = [ 
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
    .createMenu(newtonTreeConservancyMenu)
      .addItem(archiveDataForPlantingDateMenuItem, "onArchiveDataForPlantingDate")
      .addSeparator()
      .addItem(aboutThisMenuItem, "onAboutThis")
      .addToUi();
}

function onEdit(e) {
  let sheet = e.source.getActiveSheet();
  let range = sheet.getRange(plantingDateRange);

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
  let cellRange = sheet.getRange(rowIndex, sheet.getRange(groupNameRange).getColumn());
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

      let cellSuffix = streetSuffixes.find((s) => (s[0] == cellToken));

      if (cellSuffix != undefined) {
        cellToken = cellSuffix[1];
      }
    }

    cellValue += cellToken;
  });

  cellRange.setValue(cellValue);

  cellRange = sheet.getRange(rowIndex, sheet.getRange(numberOfTreeRequestedRange).getColumn());
  cellValue = cellRange.getValue();

  if (cellValue == "") {
    cellRange.setValue(0);
  }

  let applicantContactDataRanges = [
    sheet.getRange(rowIndex, sheet.getRange(firstNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(lastNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(emailAddressRange).getColumn())
  ];
  let planterContactDataRanges = [
    sheet.getRange(rowIndex, sheet.getRange(planterFirstNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(planterLastNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(planterEmailAddressRange).getColumn())
  ];

  if (applicantContactDataRanges.every((e, i) => e.getValue().toLowerCase().trim() == planterContactDataRanges[i].getValue().toLowerCase().trim())) {
    planterContactDataRanges.forEach((r) => r.setValue(""));
  }

  sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setVerticalAlignment("top");

  let emailAddress = sheet.getRange(rowIndex, sheet.getRange(emailAddressRange).getColumn()).getValue();

  if (emailAddress != undefined) {
    let senderName = sheet.getRange(applicationAckEmailSenderNameRange).getValue();
    let replyTo    = sheet.getRange(applicationAckEmailReplyToRange).getValue();
    let subject    = sheet.getRange(applicationAckEmailSubjectRange).getValue();
    let body       = sheet.getRange(applicationAckEmailBodyRange).getValue();

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

  let response = ui.prompt(archiveDataForPlantingDateTitle,
    "Enter the planting date you want to archive. Please specify \"YYYY\" followed by \"Spring\" or \"Fall\" with one space between the year and season, and the first letter of the season capitalized.\n\nExample: 2024 Spring",
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let plantingDate = response.getResponseText();

    let file  = SpreadsheetApp.getActiveSpreadsheet();
    let range = file.getRange(plantingDateRange);
    let sheet = range.getSheet();
    let hits  = range.createTextFinder(plantingDate).matchEntireCell(true).findAll();

    if (hits.length > 0) {
      let deletions = [];
      let archive   = file.insertSheet(plantingDate + " " + "Archive", file.getSheets().length);

      sheet.getRange(1, 1, 1, sheet.getLastColumn()).copyTo(archive.getRange("A1"));

      let dstRow = 2;

      hits.forEach(function(h) {
        let srcRow = h.getRow();

        sheet.getRange(srcRow, 1, 1, sheet.getLastColumn()).copyTo(archive.getRange(dstRow, 1));
        deletions.push(srcRow);
        dstRow++;
      });

      deletions.reverse().forEach(d => sheet.deleteRow(d));
    }

    ui.alert(countOfRowsArchivedTitle,
      "Number of rows archived is " + hits.length + ".",
      ui.ButtonSet.OK);
  }
}

function onAboutThis() {
  let ui = SpreadsheetApp.getUi();

  ui.alert(aboutTitle,
    "Deployment ID\n" + deploymentId + "\n\n" +
    "Version\n" + deploymentVersion + "\n\n\n" +
    "Newton Tree Conservancy\n" +
    "www.newtontreeconservancy.org",
    ui.ButtonSet.OK);
}
