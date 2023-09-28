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
  let isLegalValue = true;
  let range        = e.source.getActiveSheet().getRange(plantingDateRange);

  if (range.getLastColumn() == e.range.columnEnd) {
    let parts = e.value.trim().split(/\s+/);

    if (parts.length == 2) {
      if (!Number.isNaN(Number.parseInt(parts[0]))) {
        if ((parts[1] == "Spring") || (parts[1] == "Fall")) {
          e.range.setValue(parts[0] + " " + parts[1]);
        }
        else {
          isLegalValue = false;
        }
      }
      else if (!Number.isNaN(Number.parseInt(parts[1]))) {
        if ((parts[0] == "Spring") || (parts[0] == "Fall")) {
          e.range.setValue(parts[1] + " " + parts[0]);
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
    let ui = SpreadsheetApp.getUi();

    ui.alert("Invalid Value Specified for Planting Date",
      "Value '" + e.value + "' is invalid. The format of Planting date is \"YYYY Spring|Fall\" with one space between the year and season.",
      ui.ButtonSet.OK);

    e.range.setValue("");    
  }
}
