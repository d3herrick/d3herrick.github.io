//
// Copyright 2023 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This file includes functions to help manage and normalize data in the spreadsheet that is synchronized
// by a Google Form that residents of Newton use to submit applications for one or more trees. That sheet
// provides data to the sheet that Newton Tree Conservancy directors use to manage planting groups.
//
// @OnlyCurrentDoc
//
const plantingDateRange          = "planting_date";
const groupNameRange             = "group_name";
const firstNameRange             = "first_name";
const lastNameRange              = "last_name";
const emailAddress               = "email_address";
const planterFirstNameRange      = "planter_first_name";
const planterLastNameRange       = "planter_last_name";
const planterEmailAddress        = "planter_email_address";
const numberOfTreeRequestedRange = "number_of_trees_requested";

function onEdit(e) {
  let sheet = e.source.getActiveSheet();
  let range = sheet.getRange(plantingDateRange);

  if (sheet.getSheetId() === range.getSheet().getSheetId()) {
    let isLegalValue = true;

    if ((e.range.rowEnd > 1) && (range.getLastColumn() == e.range.columnEnd)) {
      let parts = e.value.trim().split(/\s+/);

      if (parts.length == 2) {
        if (!Number.isNaN(Number.parseInt(parts[0])) && (parts[0].length == 4)) {
          if ((parts[1] === "Spring") || (parts[1] === "Fall")) {
            e.range.setValue(parts[0] + " " + parts[1]);
          }
          else {
            isLegalValue = false;
          }
        }
        else if (!Number.isNaN(Number.parseInt(parts[1])) && (parts[1].length == 4)) {
          if ((parts[0] === "Spring") || (parts[0] === "Fall")) {
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
    replaceAll("  ", " ").
    trim();

  let cellParts = cellValue.split(" ");
  let cellIndex = 0;

  cellValue = "";

  cellParts.forEach(function(e) {
    if (cellIndex > 0) {
      cellValue += " ";  
    }

    cellValue += e.charAt(0).toUpperCase() + e.slice(1);
    cellIndex++;    
  });

  cellRange.setValue(cellValue);
  cellRange = sheet.getRange(rowIndex, sheet.getRange(numberOfTreeRequestedRange).getColumn());
  cellValue = cellRange.getValue();

  if (cellValue === "") {
    cellRange.setValue(0);
  }

  let applicantContactDataRanges = [
    sheet.getRange(rowIndex, sheet.getRange(firstNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(lastNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(emailAddress).getColumn())
  ];
  let planterContactDataRanges = [
    sheet.getRange(rowIndex, sheet.getRange(planterFirstNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(planterLastNameRange).getColumn()),
    sheet.getRange(rowIndex, sheet.getRange(planterEmailAddress).getColumn())
  ];

  if (applicantContactDataRanges.every((e, i) => e.getValue().toLowerCase() === planterContactDataRanges[i].getValue().toLowerCase())) {
    planterContactDataRanges.forEach((r) => r.setValue(""));
  }
}