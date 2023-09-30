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
// applications for one or more trees.
//
// @OnlyCurrentDoc
//
const deploymentId                   = "1WKo3XAKCpP1mwqEOKDm_IUDpv71mZsC-JiEQqnE7DKoit_OjzKUNmm6k";
const deploymentVersion              = "24";
const formDataSheetId                = "1V6U8eDIYtzxjyaP_6aifgowaJkNFcCQtGzGkDPINZ_s";
const formDataSheetRange             = "form_data";
const plantingDateRange              = "planting_date";
const groupNameRange                 = "group_name";
const groupLeaderDataFilter          = "group_leader_data_filter";
const groupDataRange                 = "group_data";
const wiresDataFilter                = "wires_data_filter";
const curbDataFilter                 = "curb_data_filter";
const plantingDataFilter             = "planting_data_filter";
const timestampDataFilter            = "timestamp_data_filter";
const lastDataRetrievalRange         = "last_data_retrieval";
const totalRequestedTreeCountRange   = "total_requested_tree_count";
const totalRecommendedTreeCountRange = "total_recommended_tree_count";
const plantingDataFilterVisibility   = "is_planting_data_filter_visible";
const newtonTreeConservancyMenu      = "Newton Tree Conservancy";
const getApplicationDataMenuItem     = "Get application data";
const toggleDataFilterMenuItem       = "Toggle data filter visibility";
const generateFileNameMenuItem       = "Generate spreadsheet file name";
const aboutThisMenuItem              = "About...";
const additionalDataAvailableTitle   = "Additional Application Data Available";
const generateFileNameTitle          = "Generate Spreadsheet File Name";
const specifyDataFilterTitle         = "Specify Application Data Filter Criteria";
const aboutTitle                     = "About Community Tree Planting Spreadsheet";
const plantingDateFilterLabel        = "Planting date";
const groupNameFilterLabel           = "Group name";

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();

  ui
    .createMenu(newtonTreeConservancyMenu)
      .addItem(getApplicationDataMenuItem, "onGetApplicationData")
      .addSeparator()
      .addItem(toggleDataFilterMenuItem, "onToggleDataFilterVisibility")
      .addSeparator()
      .addItem(generateFileNameMenuItem, "onGenerateSpreadsheetName")
      .addSeparator()
      .addItem(aboutThisMenuItem, "onAboutThis")
      .addToUi();

  let sheet = e.source.getActiveSheet();
  let rows  = listApplicationData_(sheet);

  if (!isApplicationDataEmpty_(rows)) {
    let response = ui.alert(additionalDataAvailableTitle, 
      "There " + ((rows.length > 1) ? "are" : "is") + " " + rows.length + " additional " + ((rows.length > 1) ? "applications" : "application") + " available for the " + plantingDateFilterLabel + " and " + groupNameFilterLabel + " you selected. Do you want to refresh your applicaton data now? If not, you may do so later by clicking menu item " + getApplicationDataMenuItem + ".",
      ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {
      onGetApplicationData(sheet, rows);
    }
  }
}

function onEdit(e) {
  let sheet = e.source.getActiveSheet();

  if (sheet.getRange(plantingDateRange).getA1Notation() != e.range.getA1Notation()) {
    let dataRange = sheet.getRange(groupDataRange);

    if ((dataRange.getRow() < e.range.rowStart) && (dataRange.getLastRow() > e.range.rowEnd)) {
      let needle = e.value.trim();

      if ((needle.length > 0) && (needle != "Yes") && (needle != "No")) {
        needle = needle.toLowerCase();

        let rangeNames = [groupLeaderDataFilter, wiresDataFilter, curbDataFilter];

        for (r of rangeNames) {
          let range = sheet.getRange(r);
      
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

              ui.alert("Invalid Value Specified for " + columnName,
                "Value \"" + e.value + "\" is invalid. Please specify either \"Yes\", or the letter \"Y\", or \"No\", or the letter \"N\".",
                ui.ButtonSet.OK);

              e.range.setValue("");    
            }

            break;
          }
        }
      }
    }
  }
  else {
    dataRange.setValue(groupName.getDataValidation().getCriteriaValues()[0].getValues()[0]);
  }
}

function onToggleDataFilterVisibility() {
  let sheet          = SpreadsheetApp.getActiveSheet();
  let userProperties = PropertiesService.getUserProperties();

  let isPlantingDataFilterVisible = userProperties.getProperty(plantingDataFilterVisibility);

  if ((isPlantingDataFilterVisible == null) || (isPlantingDataFilterVisible == "true")) {
    sheet.hideRow(sheet.getRange(plantingDataFilter));
    sheet.hideColumn(sheet.getRange(timestampDataFilter));
    sheet.hideColumn(sheet.getRange(groupLeaderDataFilter));

    userProperties.setProperty(plantingDataFilterVisibility, "false");
  }
  else {
    sheet.unhideRow(sheet.getRange(plantingDataFilter));
    sheet.unhideColumn(sheet.getRange(timestampDataFilter));
    sheet.unhideColumn(sheet.getRange(groupLeaderDataFilter));

    userProperties.setProperty(plantingDataFilterVisibility, "true");
  }

  sheet.setActiveSelection(sheet.getRange("A1"));
}

function onGetApplicationData(sheet, rows) {
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSheet();
  }

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

function onGenerateSpreadsheetName(sheet) {
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSheet();
  }

  let ui    = SpreadsheetApp.getUi();
  let alert = validateDataFilterCriteria_(sheet);

  if (alert.length == 0) {
    let response = ui.prompt(generateFileNameTitle,
      "Enter the first name of the director assigned to this planting group",
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
    ui.alert(specifyDataFilterTitle, 
      "Please select " + alert + " and then click menu item " + getApplicationDataMenuItem + ", and then " + generateFileNameMenuItem + ".",
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

function insertApplicationData_(sheet, rows) {
  if (rows == null) {
    rows = listApplicationData_(sheet);
  }
  
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
  let file           = sheet.getParent();
  let plantingDate   = sheet.getRange(plantingDateRange).getValue();
  let groupName      = sheet.getRange(groupNameRange).getValue();
  let dataRange      = sheet.getRange(groupDataRange);
  let firstRow       = dataRange.getRow();
  let lastRow        = dataRange.getLastRow();
  let currentRowKeys = new Set();
  let rows           = [];

  if ((lastRow - firstRow) > 1) {
    dataRange.offset(1, 0, (lastRow - firstRow - 1)).
      getValues().
      map(v => v[0].valueOf()).
      forEach(function(e) {
        currentRowKeys.add(e);
      });
  }

  // Because Google's SQL subset does not support splitting column values, and then sorting the parsed values,
  // we sort the application data ourselves. Consequently, we do not order the data explicitly in the query.
  // In addition, query results are, effectively, a (semi) sparse maxtrix; the results will include null, not
  // a zero-length string, for any number of consecutive columns that have no value on the end of a row's
  // array of values. To circumvent this, we include a constant ('$') as the last column in the query. We then
  // remove that slice of the array (and others, as necessary) before returning results from this function.
  let query = "=query(importrange(\"" + formDataSheetId + "\", \"" + formDataSheetRange + "\"), \"SELECT Col1, Col6, Col7, Col4, Col5, Col8, Col9, Col10, Col14, Col15, Col16, Col17, Col18, Col19, Col20, Col21, '$' WHERE Col1 IS NOT NULL AND lower(Col2) = lower(\'" + plantingDate + "\') AND lower(Col3) = lower(\'" + groupName + "\') label '$' ''\", 0)";

  let newSheet = file.insertSheet().hideSheet();
  newSheet.getRange(1, 1).setFormula(query);

  let newData = newSheet.getDataRange().getValues();
  file.deleteSheet(newSheet);

  if (!isApplicationDataEmpty_(newData)) {
    if (currentRowKeys.size > 0) {
      newData.forEach(function(e) {
        if (!currentRowKeys.has(e[0].valueOf())) {
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
  let op1 = d1[1].split(/\s+/);
  let op2 = d2[1].split(/\s+/);
  let s1  = op1[1].toLowerCase();
  let s2  = op2[1].toLowerCase();
  
  let rc = 0;

  if (s1 > s2) {
    rc = -1;
  }
  else if (s1 < s2) {
    rc = 1;
  }
  else {
    let t1 = op1[0].match(/\d+|\D+/g);
    let t2 = op2[0].match(/\d+|\D+/g);
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

  return rc;
}

function isApplicationDataEmpty_(rows) {
  return !((rows != null) && (rows.length > 0) && (rows[0] != "#N/A") && (rows[0] != "#VALUE!") && (rows[0] != "#ERROR!"));
}