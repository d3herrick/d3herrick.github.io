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
const deploymentId                     = "14PvqcKWB7ipcH6WytZZS4rMlmap7bnVOnGD30TgD_FIHzojPALwEzXJN";
const deploymentVersion                = "6";
const formDataSheetIdRange             = "form_data_spreadsheet_id";
const formDataSheetRange               = "form_data";
const plantingDateRange                = "planting_date";
const groupNameRange                   = "group_name";
const queryResultsRange                = "query_results"
const groupLeaderDataFilter            = "group_leader_data_filter";
const groupDataRange                   = "group_data";
const wiresDataFilter                  = "wires_data_filter";
const curbDataFilter                   = "curb_data_filter";
const largeTreeCountFilter             = "large_tree_count_filter";
const mediumTreeCountFilter            = "medium_tree_count_filter";
const smallTreeCountFilter             = "small_tree_count_filter";
const tbdTreeCountFilter               = "tbd_tree_count_filter";
const bermDataFilter                   = "berm_data_filter";
const plantingDataFilter               = "planting_data_filter";
const timestampDataFilter              = "timestamp_data_filter";
const lastDataRetrievalRange           = "last_data_retrieval";
const totalRequestedTreeCountRange     = "total_requested_tree_count";
const totalRecommendedTreeCountRange   = "total_recommended_tree_count";
const plantingDataFilterVisibility     = "is_planting_data_filter_visible";
const insertEmptyRowsMax               = 30;
const newtonTreeConservancyMenu        = "Newton Tree Conservancy";
const getApplicationDataMenuItem       = "Get application data";
const toggleDataFilterMenuItem         = "Toggle data filter visibility";
const generateFileNameMenuItem         = "Generate spreadsheet file name";
const insertEmptyRowsMenuItem          = "Insert empty rows";
const aboutThisMenuItem                = "About...";
const additionalDataAvailableTitle     = "Additional Application Data Available";
const generateFileNameTitle            = "Generate Spreadsheet File Name";
const insertEmptyRowsTitle             = "Insert Empty Rows";
const specifyDataFilterTitle           = "Specify Application Data Filter Criteria";
const specifiedInvalidColumnValueTitle = "Invalid Value Specified for ";
const aboutTitle                       = "About Community Tree Planting Spreadsheet";
const plantingDateFilterLabel          = "Planting date";
const groupNameFilterLabel             = "Group name";

const booleanValidationFilters = [
  [groupLeaderDataFilter, ""],
  [wiresDataFilter, ""],
  [curbDataFilter, ""]
];

const integerValidationFilters = [
  [largeTreeCountFilter,  ""],
  [mediumTreeCountFilter, ""],
  [smallTreeCountFilter,  ""],
  [tbdTreeCountFilter,    ""],
  [bermDataFilter,        "If there is no berm, specify a width of zero."]
];

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();

  ui
    .createMenu(newtonTreeConservancyMenu)
      .addItem(getApplicationDataMenuItem, "onGetApplicationData")
      .addSeparator()
      .addItem(toggleDataFilterMenuItem, "onToggleDataFilterVisibility")
      .addSeparator()
      .addItem(generateFileNameMenuItem, "onGenerateSpreadsheetName")
      .addItem(insertEmptyRowsMenuItem, "onInsertEmptyRows")
      .addSeparator()
      .addItem(aboutThisMenuItem, "onAboutThis")
      .addToUi();

  let sheet = getMainSheet_();

  let plantingDate     = sheet.getRange(plantingDateRange);
  let groupName        = sheet.getRange(groupNameRange);
  let archivedDataNote = null;

  if (plantingDate.getDataValidation().getCriteriaValues()[0].getValues().find((v) => (v[0] === plantingDate.getValue())) != undefined) {
    let rows = listApplicationData_(sheet);

    if (!isApplicationDataEmpty_(rows)) {
      let response = ui.alert(additionalDataAvailableTitle, 
        "There " + ((rows.length > 1) ? "are" : "is") + " " + rows.length + " additional " + ((rows.length > 1) ? "applications" : "application") + " available for the " + plantingDateFilterLabel + " and " + groupNameFilterLabel + " you selected. Do you want to refresh your applicaton data now? If not, you may do so later by clicking menu item " + getApplicationDataMenuItem + ".",
        ui.ButtonSet.YES_NO);

      if (response == ui.Button.YES) {
        onGetApplicationData(sheet, rows);
      }
    }
  }
  else {
    archivedDataNote = "The specified planting date has concluded, so data associated with the planting date has been archived."
  }

  plantingDate.setNote(archivedDataNote);
  groupName.setNote(archivedDataNote);
}

function onEdit(e) {
  let sheet = getMainSheet_();

  if (sheet.getRange(plantingDateRange).getA1Notation() != e.range.getA1Notation()) {
    let dataRange = sheet.getRange(groupDataRange);

    if ((dataRange.getRow() < e.range.rowStart) && (dataRange.getLastRow() > e.range.rowEnd)) {
      if (e.value !== undefined) {
        let needle = e.value.trim();

        if ((needle.length > 0) && (needle !== "Yes") && (needle !== "No")) {
          needle = needle.toLowerCase();

          for (r of booleanValidationFilters) {
            let range = sheet.getRange(r[0]);
        
            if (range.getLastColumn() == e.range.columnEnd) {
              if ((needle === "y") || (needle === "yes")) {
                e.range.setValue("Yes");
              }
              else if ((needle === "n") || (needle === "no")) {
                e.range.setValue("No");
              }
              else {
                let ui         = SpreadsheetApp.getUi();
                let columnName = sheet.getRange(dataRange.getRow(), range.getLastColumn()).getValue();

                ui.alert(specifiedInvalidColumnValueTitle + columnName,
                  "Value \"" + e.value + "\" is invalid. Please specify either \"Yes\", or the letter \"Y\", or \"No\", or the letter \"N\". " + r[1],
                  ui.ButtonSet.OK);

                e.range.setValue("");    
              }

              break;
            }
          }

          for (r of integerValidationFilters) {
            let range = sheet.getRange(r[0]);
        
            if (range.getLastColumn() == e.range.columnEnd) {
              let cellValue = Number.parseInt(needle);

              if (!Number.isInteger(cellValue)) {
                let ui         = SpreadsheetApp.getUi();
                let columnName = sheet.getRange(dataRange.getRow(), range.getLastColumn()).getValue();

                ui.alert(specifiedInvalidColumnValueTitle + columnName,
                  "Value \"" + e.value + "\" is invalid. Please specify an integer greater than or equal to zero. " + r[1],
                  ui.ButtonSet.OK);

                e.range.setValue("");    
              }

              break;
            }
          }
        }
      }
    }
  }
  else {
    let range = sheet.getRange(groupNameRange);

    range.setValue(range.getDataValidation().getCriteriaValues()[0].getValues()[0]);
  }
}

function onToggleDataFilterVisibility() {
  let sheet      = getMainSheet_();
  let properties = PropertiesService.getDocumentProperties();

  let isPlantingDataFilterVisible = properties.getProperty(plantingDataFilterVisibility);

  if ((isPlantingDataFilterVisible == null) || (isPlantingDataFilterVisible == "true")) {
    sheet.hideRow(sheet.getRange(plantingDataFilter));
    sheet.hideColumn(sheet.getRange(timestampDataFilter));
    sheet.hideColumn(sheet.getRange(groupLeaderDataFilter));

    properties.setProperty(plantingDataFilterVisibility, "false");
  }
  else {
    sheet.unhideRow(sheet.getRange(plantingDataFilter));
    sheet.unhideColumn(sheet.getRange(timestampDataFilter));
    sheet.unhideColumn(sheet.getRange(groupLeaderDataFilter));

    properties.setProperty(plantingDataFilterVisibility, "true");
  }

  sheet.setActiveSelection(sheet.getRange("A1"));
}

function onGetApplicationData(sheet = getMainSheet_(), rows) {
  let ui    = SpreadsheetApp.getUi();
  let alert = validateDataFilterCriteria_(sheet);

  if (alert.length == 0) {
    insertApplicationData_(sheet, rows);

    sheet.getRange(lastDataRetrievalRange).setValue(new Date());
  }
  else {
    ui.alert(specifyDataFilterTitle, 
      "Please select " + alert + " and then click menu item " + getApplicationDataMenuItem + ".",
      ui.ButtonSet.OK);
  }
}

function onGenerateSpreadsheetName() {
  let sheet = getMainSheet_();
  let ui    = SpreadsheetApp.getUi();
  let alert = validateDataFilterCriteria_(sheet);

  if (alert.length == 0) {
    let response = ui.prompt(generateFileNameTitle,
      "Enter the first name of the director assigned to this planting group:",
      ui.ButtonSet.OK_CANCEL);

    let directorName = response.getResponseText();

    if ((response.getSelectedButton() == ui.Button.OK) && (directorName.length > 0)) {
      let plantingDate   = sheet.getRange(plantingDateRange).getValue();
      let groupName      = sheet.getRange(groupNameRange).getValue();
      let totalTreeCount = sheet.getRange(totalRecommendedTreeCountRange).getValue();

      if ((totalTreeCount == null) || (totalTreeCount == 0)) {
        totalTreeCount = sheet.getRange(totalRequestedTreeCountRange).getValue();

        if (totalTreeCount == null) {
          totalTreeCount = 0;
        }
      }

      let spreadSheetName = plantingDate + "-" + groupName + " (" + directorName + ") (" + totalTreeCount + ")";

      ui.alert(generateFileNameTitle,
        "Copy and paste the name below, setting it as the name of your spreadsheet:\n\n" + spreadSheetName,
        ui.ButtonSet.OK);
    }
  }
  else {
    ui.alert(generateFileNameTitle, 
      "Please select " + alert + " and then click menu item " + getApplicationDataMenuItem + ", and then " + generateFileNameMenuItem + ".",
      ui.ButtonSet.OK);
  }
}

function onInsertEmptyRows() {
  let sheet    = getMainSheet_();
  let ui       = SpreadsheetApp.getUi();
  let rowIndex = sheet.getRange(groupDataRange).getLastRow();

  let response = ui.prompt(insertEmptyRowsTitle,
    "Enter the number of empty rows to insert. You may specify up to " + insertEmptyRowsMax + " rows. The empty rows will be inserted starting at row " + rowIndex + ":",
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let rowCount = Number.parseInt(response.getResponseText());

    if (Number.isInteger(rowCount) && (rowCount > 0)) {
      if (rowCount <= insertEmptyRowsMax) {
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
        ui.alert(insertEmptyRowsTitle,
          "You may not specify more than " + insertEmptyRowsMax + " rows.",
          ui.ButtonSet.OK);
      }
    }
    else {
      ui.alert(insertEmptyRowsTitle,
        response.getResponseText() + " is not a valid number.",
        ui.ButtonSet.OK);
    }
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

function validateDataFilterCriteria_(sheet) {
  let plantingDate          = sheet.getRange(plantingDateRange).getValue();
  let plantingDateFirstItem = sheet.getRange(plantingDateRange).getDataValidation().getCriteriaValues()[0].getValue();
  let groupName             = sheet.getRange(groupNameRange).getValue();
  let groupNameFirstItem    = sheet.getRange(groupNameRange).getDataValidation().getCriteriaValues()[0].getValue();

  let alert = "";

  if (plantingDate == plantingDateFirstItem) {
    alert += plantingDateFilterLabel;
  }

  if (groupName == groupNameFirstItem) {
    if (alert.length > 0) {
      alert += " and ";
    }

    alert += groupNameFilterLabel;
  }

  return alert;
}

function insertApplicationData_(sheet, rows = listApplicationData_(sheet)) {
  if (!isApplicationDataEmpty_(rows)) {
    let dataRange = sheet.getRange(groupDataRange);
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

        dataRange = sheet.getRange(groupDataRange);
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
  let file            = sheet.getParent();
  let plantingDate    = sheet.getRange(plantingDateRange).getValue();
  let groupName       = sheet.getRange(groupNameRange).getValue();
  let dataRange       = sheet.getRange(groupDataRange);
  let formDataSheetId = file.getRange(formDataSheetIdRange).getValue();
  let firstRow        = dataRange.getRow();
  let lastRow         = dataRange.getLastRow();
  let currentRowKeys  = new Set();
  let rows            = [];

  if ((lastRow - firstRow) > 1) {
    dataRange.offset(1, 0, (lastRow - firstRow - 1)).
      getValues().
      forEach(function(e) {
        if (e[0] !== '') {
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
  let query = "=query(importrange(\"" + formDataSheetId + "\", \"" + formDataSheetRange + "\"), \"SELECT Col1, Col6, Col7, Col4, Col5, Col8, Col9, Col10, Col14, Col15, Col16, Col17, Col18, Col19, Col20, Col21, '$' WHERE Col1 IS NOT NULL AND lower(Col2) = lower(\'" + plantingDate + "\') AND lower(Col3) = lower(\'" + groupName + "\') label '$' ''\", 0)";

  let queryResults = sheet.getRange(queryResultsRange);
  let newData      = null;

  try {
    queryResults.setFormula(query);

    newData = queryResults.getSheet().getDataRange().getValues();
  }
  finally {
    queryResults.clear();
  }

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
  let streetAddress = row[1];
  let zipcode       = row[2];

  row[1] = streetAddress + "\n" + zipcode;
  row.splice(2, 1);
}

function mergeResidentName_(row) {
  let residentFirstName = row[2];
  let residentLastName  = row[3];

  row[2] = residentFirstName + " " + residentLastName;
  row.splice(3, 1);
}

function mergeInfrastructureIssues_(row) {
  let residentNotes       = row[7];
  let pendingConstruction = row[8];
  let pendingUtilityWork  = row[9];
  let naturalGasLeaks     = row[10];
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

    switch (issueList.length) {
      case 1:
        residentNotes += " " + issueList[0];
        break;

      case 2:
        residentNotes += " " + issueList[0] + " and " + issueList[1];
        break;

      case 3:
        residentNotes += " " + issueList[0] + ", " + issueList[1] + " and " + issueList[2];
        default:
    }

    row[7] = residentNotes + ".";
  }

  row.splice(8, 3);
}

function mergePlanterContact_(row) {
  let residentNotes       = row[7];
  let planterFirstName    = row[8];
  let planterLastName     = row[9];
  let planterEmailAddress = row[10];

  if ((planterFirstName.length > 0) || (planterLastName.length > 0) || (planterEmailAddress.length > 0)) {
    if (residentNotes.length > 0) {
        residentNotes += "\n\n";
    }

    residentNotes += "Designated planter is";
    
    if (planterFirstName.length > 0) {
      residentNotes += " " + planterFirstName;
    }

    if (planterLastName.length > 0) {
      residentNotes += " " + planterLastName;
    }

    if (planterEmailAddress.length > 0) {
      residentNotes += " at " + planterEmailAddress;
    }

    row[7] = residentNotes + ".";
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
        if (tokens.charAt(i + 1) === ' ') {
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
    else if (c === ' ') {
      if (i < (length - 2)) {
        if ((/^[a-z]/.test(tokens.charAt(i + 1)) && (tokens.charAt(i + 2) === ' '))) {
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

    if ((c === ' ') || (c === '-')) {
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
    let bits  = token.split(" ");

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

function getMainSheet_() {
  let sheet = undefined;
  let range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(groupDataRange);

  if (range != null) {
    sheet = range.getSheet();
  }
  else {
    throw new Error("Main sheet could not be resolved");
  }

  return sheet;
}

function isApplicationDataEmpty_(rows) {
  return !((rows != null) && (rows.length > 0) && (rows[0] != "#N/A") && (rows[0] != "#VALUE!") && (rows[0] != "#ERROR!"));
}
