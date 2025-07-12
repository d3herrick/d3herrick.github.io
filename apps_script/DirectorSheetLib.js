//
// Copyright 2023 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This library includes functions for the spreadsheets that directors of the Newton Tree Conservancy
// use to manage neighborhood tree planting groups. The Google Sheet that provides data to the tree
// planting group sheets is synchronized by a Google Form that residents of Newton use to submit
// applications for trees to be planted on their property.
//
// @OnlyCurrentDoc
//
const DEPLOYMENT_ID                          = "14PvqcKWB7ipcH6WytZZS4rMlmap7bnVOnGD30TgD_FIHzojPALwEzXJN";
const DEPLOYMENT_VERSION                     = "37";
const FORM_DATA_SHEET_ID_RANGE               = "form_data_spreadsheet_id";
const FORM_DATA_SHEET_RANGE                  = "form_data";
const PLANTING_DATE_RANGE                    = "planting_date";
const GROUP_NAME_RANGE                       = "group_name";
const GROUP_LEADER_DATA_FILTER               = "group_leader_data_filter";
const PHONE_DATA_FILTER                      = "phone_data_filter";
const NUMBER_OF_TREES_REQUESTED_FILTER       = "number_of_trees_requested";
const RESIDENT_NOTES_RANGE                   = "resident_notes"
const GROUP_DATA_RANGE                       = "group_data";
const RECOMMENDED_TREE_COUNT_DATA_RANGE      = "recommended_tree_count_data"
const LARGE_TREE_COUNT_FILTER                = "large_tree_count_filter";
const MEDIUM_TREE_COUNT_FILTER               = "medium_tree_count_filter";
const SMALL_TREE_COUNT_FILTER                = "small_tree_count_filter";
const TBD_TREE_COUNT_FILTER                  = "tbd_tree_count_filter";
const WIRES_DATA_FILTER                      = "wires_data_filter";
const CURB_DATA_FILTER                       = "curb_data_filter";
const BERM_DATA_FILTER                       = "berm_data_filter";
const NTC_NOTES_RANGE                        = "ntc_notes";
const GAS_LEAK_TESTING_NOTES_FILTER          = "gas_leak_testing_notes";
const PLANTING_DATA_FILTER                   = "planting_data_filter";
const TIMESTAMP_DATA_FILTER                  = "timestamp_data_filter";
const LAST_DATA_RETRIEVAL_RANGE              = "last_data_retrieval";
const TOTAL_REQUESTED_TREE_COUNT_RANGE       = "total_requested_tree_count";
const TOTAL_RECOMMENDED_TREE_COUNT_RANGE     = "total_recommended_tree_count";
const PLANTING_DATA_FILTER_VISIBILITY        = "is_planting_data_filter_visible";
const INSERT_EMPTRY_ROWS_MAX                 = 30;
const DIRECTOR_NAME_PROP                     = "director_name_prop";
const DIRECTOR_NAME_NOT_SPECIFIED            = "Not specified";
const DUPLICATE_ROW_COLOR                    = "darkgray";
const DUPLICATE_ROW_REQUESTED_VALUE          = "X";
const DUPLICATE_ROW_REQUESTED_FONT_SIZE      = 45;
const RETRIEVING_DATA_STATUS                 = "Retrieving data..."
const NEWTON_TREE_CONSERVANCY_MENU           = "Newton Tree Conservancy";
const GET_APPLICATION_DATA_MENU_ITEM         = "Get application data";
const TOGGLE_DATA_FILTER_MENU_ITEM           = "Toggle data filter visibility";
const SET_DIRECTOR_FILE_NAME_MENU_ITEM       = "Set director for spreadsheet file name";
const DUPLICATE_ROW_FOR_CORNER_LOT_MENU_ITEM = "Duplicate application row for corner lot"
const INSERT_EMPTY_ROWS_MENU_ITEM            = "Insert empty rows";
const ABOUT_MENU_ITEM                        = "About...";
const ADDITIONAL_DATA_AVAILABLE_TITLE        = "Additional Application Data Available";
const SET_DIRECTOR_FILE_NAME_TITLE           = "Set Director for Spreadsheet File Name";
const DUPLICATE_ROW_FOR_CORNER_LOT_TITLE     = "Duplicate Application Row for Corner Lot"
const INSERT_EMPTY_ROWS_TITLE                = "Insert Empty Rows";
const SPECIFY_DATA_FILTER_TITLE              = "Specify Application Data Filter Criteria";
const SPECIFIED_INVALID_COLUMN_VALUE_TITLE   = "Invalid Value Specified";
const ABOUT_TITLE                            = "About Community Tree Planting Spreadsheet";
const PLANTING_DATA_FILTER_LABEL             = "Planting date";
const GROUP_NAME_FILTER_LABEL                = "Group name";
const ARCHIVED_DATA_NOTE                     = "Because it has concluded, data associated with the planting date has been archived.";
const CORNER_LOT_STREET_TAG                  = "replace-with-street-name: "
const CORNER_LOT_RESIDENT_NOTE_TAG           = "Application row above was duplicated to accommodate planting on multiple streets that border the corner lot.";

const BOOLEAN_VALIDATION_FILTERS = [
  [GROUP_LEADER_DATA_FILTER, ""],
  [WIRES_DATA_FILTER, ""],
  [CURB_DATA_FILTER, ""]
];

const INTEGRER_VALIDATION_FILTERS = [
  [LARGE_TREE_COUNT_FILTER,  ""],
  [MEDIUM_TREE_COUNT_FILTER, ""],
  [SMALL_TREE_COUNT_FILTER,  ""],
  [TBD_TREE_COUNT_FILTER,    ""],
  [BERM_DATA_FILTER,         "If there is no berm, specify a width of zero."]
];

function onOpen(e) {
  let sheet = getGroupDataSheet_();
  let ui    = SpreadsheetApp.getUi();

  ui
    .createMenu(NEWTON_TREE_CONSERVANCY_MENU)
      .addItem(GET_APPLICATION_DATA_MENU_ITEM, "onGetApplicationData")
      .addItem(SET_DIRECTOR_FILE_NAME_MENU_ITEM, "onSetDirectorFileName")
      .addSeparator()
      .addItem(DUPLICATE_ROW_FOR_CORNER_LOT_MENU_ITEM, "onDuplicateRowForCornerLot")
      .addItem(INSERT_EMPTY_ROWS_MENU_ITEM, "onInsertEmptyRows")
      .addItem(TOGGLE_DATA_FILTER_MENU_ITEM, "onToggleDataFilterVisibility")
      .addSeparator()
      .addItem(ABOUT_MENU_ITEM, "onAbout")
      .addToUi();

  let plantingDate   = sheet.getRange(PLANTING_DATE_RANGE);
  let groupName      = sheet.getRange(GROUP_NAME_RANGE);
  let isDataArchived = true;

  if (hasDataFilterValidators_()) { 
    let plantingDateValidation = plantingDate.getDataValidation();

    if (plantingDateValidation.getCriteriaValues()[0].getValues().find((v) => (v[0] == plantingDate.getValue())) != undefined) {
      let rows = listApplicationData_(sheet);

      if (!isApplicationDataEmpty_(rows)) {
        let response = ui.alert(ADDITIONAL_DATA_AVAILABLE_TITLE, 
          `There ${((rows.length > 1) ? "are" : "is")} ${rows.length} additional ${((rows.length > 1) ? "applications" : "application")} available for the ${PLANTING_DATA_FILTER_LABEL} and  ${GROUP_NAME_FILTER_LABEL} you selected. Do you want to refresh your applicaton data now? If not, you may do so later by clicking menu item ${GET_APPLICATION_DATA_MENU_ITEM}.`,
          ui.ButtonSet.YES_NO);

        if (response == ui.Button.YES) {
          onGetApplicationData(rows);
        }
      }

      isDataArchived = false;
    }
    else {
      plantingDate.setDataValidation(null);
      plantingDate.protect().setWarningOnly(true);
      groupName.setDataValidation(null);
      groupName.protect().setWarningOnly(true);
    }
  }

  if (isDataArchived) {
    plantingDate.setNote(ARCHIVED_DATA_NOTE);
    groupName.setNote(ARCHIVED_DATA_NOTE);
  }
  else {
    setSpreadsheetFileName_();
  }
}

function onEdit(e) {
  let sheet = getGroupDataSheet_();

  if (sheet.getRange(PLANTING_DATE_RANGE).getA1Notation() != e.range.getA1Notation()) {
    let dataRange = sheet.getRange(GROUP_DATA_RANGE);

    if ((dataRange.getRow() < e.range.rowStart) && (dataRange.getLastRow() > e.range.rowEnd)) {
      let isValidValue = true;

      if (e.value != undefined) {
        let needle = e.value.trim();

        if (needle.length > 0) {
          needle = needle.toLowerCase();

          for (r of BOOLEAN_VALIDATION_FILTERS) {
            let range = sheet.getRange(r[0]);

            if (range.getLastColumn() == e.range.columnEnd) {
              if ((needle == "y") || (needle == "yes")) {
                e.range.setValue("Yes");
              }
              else if ((needle == "n") || (needle == "no")) {
                e.range.setValue("No");
              }
              else {
                let ui         = SpreadsheetApp.getUi();
                let columnName = sheet.getRange(dataRange.getRow(), range.getLastColumn()).getValue();

                ui.alert(`${SPECIFIED_INVALID_COLUMN_VALUE_TITLE} for ${columnName}`,
                  `Value "${e.value}" is invalid. Please specify either "Yes", or the letter "Y", or "No", or the letter "N". ${r[1]}`,
                  ui.ButtonSet.OK);

                e.range.setValue("");    
                isValidValue = false;

                break;
              }
            }
          }

          if (isValidValue) {
            for (r of INTEGRER_VALIDATION_FILTERS) {
              let range = sheet.getRange(r[0]);

              if (range.getLastColumn() == e.range.columnEnd) {
                if (!Number.isInteger(Number(needle))) {
                  let ui         = SpreadsheetApp.getUi();
                  let columnName = sheet.getRange(dataRange.getRow(), range.getLastColumn()).getValue();

                  ui.alert(`${SPECIFIED_INVALID_COLUMN_VALUE_TITLE} for ${columnName}`,
                    `Value "${e.value}" is invalid. Please specify an integer greater than or equal to zero. ${r[1]}`,
                    ui.ButtonSet.OK);

                  e.range.setValue("");    
                  isValidValue = false;

                  break;
                }
              }
            }
          }
        }
      }
      else if (e.oldValue != undefined) {
        isValidValue = true;
      }

      if (isValidValue) {
        dataRange = sheet.getRange(RECOMMENDED_TREE_COUNT_DATA_RANGE);

        if ((dataRange.getRow() < e.range.rowStart) && (dataRange.getLastRow() > e.range.rowEnd)) {
          setSpreadsheetFileName_();
        }
      }
    }
  }
  else {
    if (hasDataFilterValidators_()) { 
      let range          = sheet.getRange(GROUP_NAME_RANGE);
      let dataValidation = range.getDataValidation();

      range.setValue(dataValidation.getCriteriaValues()[0].getValues()[0]);
    }
  }
}

function onGetApplicationData(rows) {
  let sheet    = getGroupDataSheet_();
  let ui       = SpreadsheetApp.getUi();
  let criteria = validateDataFilterCriteria_();

  if (criteria.isComplete) {
    let range  = sheet.getRange(LAST_DATA_RETRIEVAL_RANGE);
    let format = range.getNumberFormat(); 

    try {
      range.setValue(RETRIEVING_DATA_STATUS);
      insertApplicationData_(sheet, rows);
    }
    finally {
      range.setValue(new Date()).setNumberFormat(format);
    }
  }
  else {
    ui.alert(SPECIFY_DATA_FILTER_TITLE, 
      `Please select ${criteria.message} and then click menu item ${GET_APPLICATION_DATA_MENU_ITEM} again.`,
      ui.ButtonSet.OK);
  }
}

function onSetDirectorFileName() {
  let sheet    = getGroupDataSheet_();
  let ui       = SpreadsheetApp.getUi();
  let criteria = validateDataFilterCriteria_();

  if (criteria.isComplete) {
    let response = ui.prompt(SET_DIRECTOR_FILE_NAME_TITLE,
      `Enter the first name of the director assigned to this planting group. If multiple directors are assigned, separate their first names with "and":`,
      ui.ButtonSet.OK_CANCEL);

    let directorName = response.getResponseText();

    if (response.getSelectedButton() == ui.Button.OK) {
      if (directorName.length > 0) {
        PropertiesService.getDocumentProperties().setProperty(DIRECTOR_NAME_PROP, directorName);

        setSpreadsheetFileName_();
      }
      else {
        PropertiesService.getDocumentProperties().deleteProperty(DIRECTOR_NAME_PROP);

        ui.alert(SET_DIRECTOR_FILE_NAME_TITLE,
          `You did not specify the name of a director. Consequently, automatic update of the spreadsheet file name will be disabled.`,
          ui.ButtonSet.OK);
      }
    }
  }
  else {
    ui.alert(SET_DIRECTOR_FILE_NAME_TITLE, 
      `Please select ${criteria.message} and then click menu item ${SET_DIRECTOR_FILE_NAME_MENU_ITEM} again.`,
      ui.ButtonSet.OK);
  }
}

function onDuplicateRowForCornerLot() {
  let sheet     = getGroupDataSheet_();
  let ui        = SpreadsheetApp.getUi();
  let row       = sheet.getActiveCell().getRow();
  let dataRange = sheet.getRange(GROUP_DATA_RANGE);

  if ((dataRange.getRow() < row) && (dataRange.getLastRow() > row)) {
    let column = sheet.getRange(GROUP_LEADER_DATA_FILTER).getLastColumn();
    let newRow = row + 1;
    let range  = sheet.getRange(row, 1, 1, column);

    sheet.insertRowsAfter(row, 1);
    range.copyTo(sheet.getRange(newRow, 1, 1, column));
    sheet.getRange(newRow, 1, 1, column).setFontColor(DUPLICATE_ROW_COLOR);

    range = sheet.getRange(newRow, sheet.getRange(NUMBER_OF_TREES_REQUESTED_FILTER).getLastColumn());
    range.setValue(DUPLICATE_ROW_REQUESTED_VALUE).setFontSize(DUPLICATE_ROW_REQUESTED_FONT_SIZE).setFontColor(DUPLICATE_ROW_COLOR);

    range = sheet.getRange(row, sheet.getRange(NTC_NOTES_RANGE).getLastColumn());
    range.setValue(CORNER_LOT_STREET_TAG + range.getValue());

    range = sheet.getRange(newRow, sheet.getRange(RESIDENT_NOTES_RANGE).getLastColumn());
    range.setValue(CORNER_LOT_RESIDENT_NOTE_TAG).setFontStyle("italic").setFontColor(DUPLICATE_ROW_COLOR);

    range = sheet.getRange(newRow, sheet.getRange(NTC_NOTES_RANGE).getLastColumn());
    range.setValue(CORNER_LOT_STREET_TAG);
  }
  else {
    ui.alert(DUPLICATE_ROW_FOR_CORNER_LOT_TITLE, 
      `Please select one of the applications located in rows ${(dataRange.getRow() + 1)} through  ${(dataRange.getLastRow() - 1)}.`,
      ui.ButtonSet.OK);
  }
}

function onInsertEmptyRows() {
  let sheet    = getGroupDataSheet_();
  let ui       = SpreadsheetApp.getUi();
  let rowIndex = sheet.getRange(GROUP_DATA_RANGE).getLastRow();

  let response = ui.prompt(INSERT_EMPTY_ROWS_TITLE,
    `Enter the number of empty rows to insert. You may specify up to ${INSERT_EMPTRY_ROWS_MAX} rows. The empty rows will be inserted starting at row ${rowIndex}:`,
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let rowCount = Number.parseInt(response.getResponseText());

    if (Number.isInteger(rowCount) && (rowCount > 0)) {
      if (rowCount <= INSERT_EMPTRY_ROWS_MAX) {
        let rowTimestamp = new Date();

        while (true) {
          if (rowCount-- > 0) {
            sheet.insertRowsBefore(rowIndex, 1);
            sheet.getRange(rowIndex, 1).setValue(rowTimestamp);

            rowIndex++;
            rowTimestamp.setSeconds(rowTimestamp.getSeconds() + 1);
          }
          else {
            break;
          }
        }
      }
      else {
        ui.alert(INSERT_EMPTY_ROWS_TITLE,
          `You may not specify more than ${INSERT_EMPTRY_ROWS_MAX} rows.`,
          ui.ButtonSet.OK);
      }
    }
    else {
      ui.alert(INSERT_EMPTY_ROWS_TITLE,
        `${response.getResponseText()} is not a valid number.`,
        ui.ButtonSet.OK);
    }
  }
}

function onToggleDataFilterVisibility() {
  let sheet      = getGroupDataSheet_();
  let properties = PropertiesService.getDocumentProperties();

  let isPlantingDataFilterVisible = properties.getProperty(PLANTING_DATA_FILTER_VISIBILITY);

  if ((isPlantingDataFilterVisible == null) || (isPlantingDataFilterVisible == "true")) {
    sheet.hideRow(sheet.getRange(PLANTING_DATA_FILTER));
    sheet.hideColumn(sheet.getRange(TIMESTAMP_DATA_FILTER));
    sheet.hideColumn(sheet.getRange(GROUP_LEADER_DATA_FILTER));
    sheet.hideColumn(sheet.getRange(PHONE_DATA_FILTER));
    sheet.hideColumn(sheet.getRange(GAS_LEAK_TESTING_NOTES_FILTER));

    properties.setProperty(PLANTING_DATA_FILTER_VISIBILITY, "false");
  }
  else {
    sheet.unhideRow(sheet.getRange(PLANTING_DATA_FILTER));
    sheet.unhideColumn(sheet.getRange(TIMESTAMP_DATA_FILTER));
    sheet.unhideColumn(sheet.getRange(GROUP_LEADER_DATA_FILTER));
    sheet.unhideColumn(sheet.getRange(PHONE_DATA_FILTER));
    sheet.unhideColumn(sheet.getRange(GAS_LEAK_TESTING_NOTES_FILTER));

    properties.setProperty(PLANTING_DATA_FILTER_VISIBILITY, "true");
  }

  sheet.setActiveSelection(sheet.getRange("A1"));
}

function onAbout() {
  let ui           = SpreadsheetApp.getUi();
  let directorName = PropertiesService.getDocumentProperties().getProperty(DIRECTOR_NAME_PROP) ?? DIRECTOR_NAME_NOT_SPECIFIED;

  ui.alert(ABOUT_TITLE,
    `Deployment ID
    ${DEPLOYMENT_ID}

    Version
    ${DEPLOYMENT_VERSION}

    Director name
    ${directorName}


    Newton Tree Conservancy
    www.newtontreeconservancy.org`,
    ui.ButtonSet.OK);
}

function onApplyUpdates() {
  // insert any code-driven updates here
}

function setSpreadsheetFileName_() {
  let criteria = validateDataFilterCriteria_();

  if (criteria.isComplete) {
    let sheet          = getGroupDataSheet_();
    let plantingDate   = sheet.getRange(PLANTING_DATE_RANGE).getValue();
    let groupName      = sheet.getRange(GROUP_NAME_RANGE).getValue();
    let totalTreeCount = sheet.getRange(TOTAL_RECOMMENDED_TREE_COUNT_RANGE).getValue() ?? 0;
    let directorName   = PropertiesService.getDocumentProperties().getProperty(DIRECTOR_NAME_PROP);

    if (((plantingDate != null) && (plantingDate.trim().length > 0)) &&
        ((groupName != null)    && (groupName.trim().length > 0)) &&
        ((directorName != null) && (directorName.trim().length > 0)))
    {
      let spreadSheetName = `${plantingDate}-${groupName} (${directorName}) (${totalTreeCount})`;

      SpreadsheetApp.getActiveSpreadsheet().rename(spreadSheetName);
    }
  }
}

function validateDataFilterCriteria_() {
  let isComplete = true;
  let message    = "";

  if (hasDataFilterValidators_()) {
    let sheet                  = getGroupDataSheet_();
    let plantingDateValidation = sheet.getRange(PLANTING_DATE_RANGE).getDataValidation();
    let groupNameValidation    = sheet.getRange(GROUP_NAME_RANGE).getDataValidation();
    let plantingDate           = sheet.getRange(PLANTING_DATE_RANGE).getValue();
    let plantingDateFirstItem  = plantingDateValidation.getCriteriaValues()[0].getValue();
    let groupName              = sheet.getRange(GROUP_NAME_RANGE).getValue();
    let groupNameFirstItem     = groupNameValidation.getCriteriaValues()[0].getValue();

    if (plantingDate == plantingDateFirstItem) {
      isComplete = false;

      message += PLANTING_DATA_FILTER_LABEL;
    }

    if (groupName == groupNameFirstItem) {
      isComplete = false;

      if (message.length > 0) {
        message += " and ";
      }

      message += GROUP_NAME_FILTER_LABEL;
    }
  }

  return {isComplete: isComplete, message: message};
}

function insertApplicationData_(sheet, rows = listApplicationData_(sheet)) {
  if (!isApplicationDataEmpty_(rows)) {
    let dataRange = sheet.getRange(GROUP_DATA_RANGE);
    let firstRow  = dataRange.getRow();
    let lastRow   = dataRange.getLastRow();

    if ((lastRow - firstRow) > 1) {
      rows.forEach(function(e) {
        let insertionRowIndex = lastRow;

        for (let i = firstRow + 1; i < lastRow; i++) {
          let values = sheet.getRange(i, 1, 1, e.length).getValues();

          let rc = compareApplicationData_(e, values[0]);

          if (rc < 0) {
            insertionRowIndex = i;
            break;
          }
        }

        sheet.insertRowsBefore(insertionRowIndex, 1);

        let range = sheet.getRange(insertionRowIndex, 1, 1, e.length);

        range.setValues([e]);
        range.setVerticalAlignment("top");

        dataRange = sheet.getRange(GROUP_DATA_RANGE);
        firstRow  = dataRange.getRow();
        lastRow   = dataRange.getLastRow();
      });
    }
    else {
      sortApplicationData_(rows);

      sheet.insertRowsBefore(lastRow, rows.length);

      let range = sheet.getRange((firstRow + 1), 1, rows.length, rows[0].length);

      range.setValues(rows);
      range.setVerticalAlignment("top");
    }
  }
}

function listApplicationData_(sheet) {
  let spreadsheet     = sheet.getParent();
  let plantingDate    = sheet.getRange(PLANTING_DATE_RANGE).getValue();
  let groupName       = sheet.getRange(GROUP_NAME_RANGE).getValue();
  let dataRange       = sheet.getRange(GROUP_DATA_RANGE);
  let formDataSheetId = spreadsheet.getRange(FORM_DATA_SHEET_ID_RANGE).getValue();
  let firstRow        = dataRange.getRow();
  let lastRow         = dataRange.getLastRow();
  let currentRowKeys  = new Set();
  let rows            = [];

  if ((lastRow - firstRow) > 1) {
    dataRange.offset(1, 0, (lastRow - firstRow - 1)).
      getValues().
      forEach(function(e) {
        if (e[0] != '') {
          currentRowKeys.add(e[0].getTime());
        }
      });
  }

  // Because Google's SQL subset does not support splitting column values, and then sorting the parsed values,
  // we sort the application data ourselves. Consequently, we do not order the data explicitly in the query.
  // In addition, query results are, effectively, a (semi) sparse maxtrix; the results will include null, not
  // a zero-length string, for any number of consecutive columns that have no value on the end of a row's
  // array of values. To circumvent this, we include a constant ('$') as the last column in the query. We then
  // remove that slice of the array (and others, as necessary) before returning results from this function.
  let query = `=query(importrange("${formDataSheetId}", "${FORM_DATA_SHEET_RANGE}"), "SELECT Col1, Col6, Col7, Col4, Col5, Col8, Col9, Col10, Col14, Col15, Col16, Col17, Col18, Col19, Col20, Col21, '$' WHERE Col1 IS NOT NULL AND lower(Col2) = lower('${plantingDate}') AND lower(Col3) = lower('${groupName}') label '$' ''", 0)`;

  let newData = executeQuery_(spreadsheet, query);

  if (!isApplicationDataEmpty_(newData)) {
    if (currentRowKeys.size > 0) {
      newData.forEach(function(e) {
        if (!currentRowKeys.has(e[0].getTime())) {
          rows.push(e);
        }
      });
    }
    else {
      rows = newData;
    }

    if (!isApplicationDataEmpty_(rows)) {
      rows.forEach(function(e) {
        mergeZipcode_(e);
        mergeResidentName_(e);
        mergeInfrastructureIssues_(e);
        mergePlanterContact_(e);
      });
    }
  }

  return rows;
}

function mergeZipcode_(row) {
  let streetAddress = row[1].trim();
  let zipcode       = row[2].trim();

  row[1] = `${streetAddress}\n${zipcode}`;
  row.splice(2, 1);
}

function mergeResidentName_(row) {
  let residentFirstName = row[2].trim();
  let residentLastName  = row[3].trim();

  row[2] = `${residentFirstName} ${residentLastName}`;
  row.splice(3, 1);
}

function mergeInfrastructureIssues_(row) {
  let residentNotes       = row[7].trim();
  let pendingConstruction = row[8].trim();
  let pendingUtilityWork  = row[9].trim();
  let naturalGasLeaks     = row[10].trim();
  let issueList           = [];

  if ((pendingConstruction == "Yes") || (pendingUtilityWork == "Yes") ||  (naturalGasLeaks == "Yes")) {
    if (residentNotes.length > 0) {
        residentNotes += "\n\n";
    }

    residentNotes += "Possible infrastructure issues include";

    if (pendingConstruction == "Yes") {
      issueList.push("pending street construction");
    }

    if (pendingUtilityWork == "Yes") {
      issueList.push("pending utility work");
    }

    if (naturalGasLeaks == "Yes") {
      issueList.push("natural gas leaks");
    }

    row[7] = `${residentNotes} ${new Intl.ListFormat('en-GB', {style: 'long', type: 'conjunction'}).format(issueList)}.`;
  }

  row.splice(8, 3);
}

function mergePlanterContact_(row) {
  let residentNotes       = row[7].trim();
  let planterFirstName    = row[8].trim();
  let planterLastName     = row[9].trim();
  let planterEmailAddress = row[10].trim();

  if ((planterFirstName.length > 0) || (planterLastName.length > 0) || (planterEmailAddress.length > 0)) {
    if (residentNotes.length > 0) {
        residentNotes += "\n\n";
    }

    residentNotes += "Designated planter is";

    if (planterFirstName.length > 0) {
      residentNotes += ` ${planterFirstName}`;
    }

    if (planterLastName.length > 0) {
      residentNotes += ` ${planterLastName}`;
    }

    if (planterEmailAddress.length > 0) {
      residentNotes += ` at ${planterEmailAddress}`;
    }

    row[7] = `${residentNotes}.`;
  }

  row.splice(8, 4);
}

function sortApplicationData_(rows) {
  if (!isApplicationDataEmpty_(rows) && (rows.length > 1)) {
    rows.sort(function(a, b) {
      return compareApplicationData_(a, b);
    });
  }

  return rows;
}

function compareApplicationData_(d1, d2) {
  let rc = 0;

  let op1 = parseStreetAddress_(d1[1]);
  let op2 = parseStreetAddress_(d2[1])

  if (op1.length > 1) {
    if (op2.length > 1) {
      let s1 = op1[1];
      let s2 = op2[1];

      if (s1 > s2) {
        rc = 1;
      }
      else if (s1 < s2) {
        rc = -1;
      }
      else {
        let t1 = op1[0].match(/\d+|\D+/);
        let t2 = op2[0].match(/\d+|\D+/);
        let n1 = Number.parseInt(t1[0]);
        let n2 = Number.parseInt(t2[0]);

        rc = (Number.isInteger(n1) ? n1 : 0) - (Number.isInteger(n2) ? n2 : 0);

        if (rc == 0) {
          if (t1.length > 1) {
            if (t2.length > 1) {
              rc = t1[1].localeCompare(t2[1]);
            }
            else {
              rc = 1;
            }
          }
          else if (t2.length > 1 ) {
            rc = -1;  
          }

          if (rc == 0) {
            rc = op1[op1.length - 1].localeCompare(op2[op2.length - 1]);
          }
        }
      }
    }
    else {
      rc = -1;
    }
  }
  else if (op2.length > 1) {
    rc = 1;
  }

  return rc;
}

function parseStreetAddress_(streetAddress) {
  let parts = [];

  let tokens = normalizeStreetAddress_(streetAddress);
  let apt    = "";
  let start  = 0;
  let end    = start;
  let length = tokens.length;

  for (let i = end; i < length; i++) {
    let c = tokens.charAt(i);

    if (Number.isInteger(Number.parseInt(c))) {
      end++;
    }
    else if ((/^[a-z]/.test(c))) {
      if (i < (length - 1)) {
        if (tokens.charAt(i + 1) == ' ') {
          end++;
          break;
        }
        else {
          break;
        }
      }
      else {
        break;
      }
    }
    else if (c == ' ') {
      if (i < (length - 2)) {
        if ((/^[a-z]/.test(tokens.charAt(i + 1)) && (tokens.charAt(i + 2) == ' '))) {
          apt = tokens.charAt(i + 1);
          end += 2;
          break;
        }
        else {
          break;
        }
      }
      else {
        break;
      }
    }
    else {
      break;
    }
  }

  if (end > start) {
    parts.push(((tokens.substring(start, (end - apt.length)).trim()) + apt).trim());
  }
  else {
    parts.push("0");
  }

  start = end;
  end   = start;

  for (let i = end; i < length; i++) {
    let c = tokens.charAt(i);

    if ((c == ' ') || (c == '-')) {
      end++;
    }
    else if (Number.isInteger(Number.parseInt(c))) {
      end++;
    }
    else {
      break;
    }
  }

  if (end >= start) {
    start = end;
    end   = length - 1;
  }

  for (let i = end; i > start; i--) {
    let c = tokens.charAt(i);

    if (Number.isInteger(Number.parseInt(c))) {
      end--;
    }
    else {
      break;
    }
  }

  if (end > start) {
    let token = tokens.substring(start, end);
    let bits  = token.split(/\s+/);

    if (bits.length > 1) {
      token = bits[0];

      for (let i = 1; i < bits.length; i++) {
        token += " " + bits[i].charAt(0);
      }
    }

    parts.push(token.trim());
    parts.push((tokens.substring(end + 1)).trim());
  }
  else {
    parts.push("");
    parts.push("00000");
  }

  return parts;
}

function normalizeStreetAddress_(streetAddress) {
  let tokens = streetAddress.toLowerCase();

  [[" - ", "-"], ["- ", "-"], [" -", "-"], [".", ""], ["  ", " "]].forEach(function(e) {
    while (tokens.indexOf(e[0]) != -1) {
      tokens = tokens.replaceAll(e[0], e[1]);
    }
  });

  let start = tokens.indexOf(",");

  if (start != -1) {
    let end    = start + 1;
    let length = tokens.length;

    for (let i = end; i < length; i++) {
      let c = tokens.charAt(i);

      if (!Number.isInteger(Number.parseInt(c))) {
        end++;
      }
      else {
        break;
      }
    }

    if (end > start) {
      tokens = tokens.replace(tokens.slice(start, end - 1), "");
    }
  }

  return tokens;
}

function hasDataFilterValidators_() {
  let sheet         = getGroupDataSheet_();
  let hasValidators = (sheet.getRange(PLANTING_DATE_RANGE).getDataValidation() != null);

  if (hasValidators) {
    hasValidators = (sheet.getRange(GROUP_NAME_RANGE).getDataValidation() != null);
  }

  return hasValidators;
}

function executeQuery_(spreadsheet, query) {
  let queryResults = spreadsheet.insertSheet().hideSheet();
  let resultData   = null;

  try {
    queryResults.getRange(1,1).setFormula(query);

    resultData = queryResults.getDataRange().getValues();
  }
  finally {
    spreadsheet.deleteSheet(queryResults);
  }

  return !isApplicationDataEmpty_(resultData) ? resultData : [];
}

function isApplicationDataEmpty_(rows) {
  return !((rows != null) && (rows.length > 0) && (rows[0] != "#N/A") && (rows[0] != "#VALUE!") && (rows[0] != "#ERROR!"));
}

function getGroupDataSheet_() {
  let sheet = undefined;
  let range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(GROUP_DATA_RANGE);

  if (range != null) {
    sheet = range.getSheet();
  }
  else {
    throw new Error("Group data sheet could not be resolved");
  }

  return sheet;
}
