//
// Copyright 2023 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This file includes functions to help manage data in the spreadsheet that is synchronized by a Google Form
// that residents of Newton use to submit applications for one or more trees. That sheet provides data to
// the sheet that Newton Tree Conservancy directors use to manage planting groups.
//
// @OnlyCurrentDoc
//
const plantingDateRange = "planting_date";

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
