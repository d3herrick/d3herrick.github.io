//
// Copyright 2025 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This library includes functions to handle import of files containing donations made to the Newton Tree
// Conservancy. These functions exist primarily to normalize and store incoming data to the layout captured
// in the spreadsheets used to manage donations and memberships.
//
// @OnlyCurrentDoc
//
const DEPLOYMENT_ID                           = "1cXoHvwTUh5pTV3_0YHl9jZsL4YZ7Ie6juG307YwOBxGLjeF81khFYHcy";
const DEPLOYMENT_VERSION                      = "1";
const DONATION_DATA_RANGE                     = "donation_data";
const PENDING_FOLDER_RANGE                    = "pending_folder";
const IMPORTED_FOLDER_RANGE                   = "imported_folder"
const FIRST_DATA_ROW_RANGE                    = "first_data_row";
const DONATION_NEW_ROW_BACKGROUND_COLOR       = "#fffcd3";
const ADDRESS_JOIN_SEPARATOR                  = ", ";
const PAYPAL_FILE_PREFIX                      = "paypal";
const PAYPAL_FILE_FIELD_COUNT                 = 41;
const PAYPAL_DONATION_PAYMENT                 = "Donation Payment"
const PAYPAL_SUBSCRIPTION_PAYMENT             = "Subscription Payment";
const PAYPAL_MOBILE_PAYMENT                   = "Mobile Payment";
const PAYPAL_MASS_PAYMENT                     = "Mass Pay Payment";
const NEWTON_TREE_CONSERVANCY_MENU            = "Newton Tree Conservancy";
const ABOUT_MENU_ITEM                         = "About...";
const IMPORT_PENDING_DONATION_DATA_MENU_ITEM  = "Import pending donation data";
const IMPORT_PENDING_DONATION_DATA_TITLE      = "Import Pending Donation Data";
const ABOUT_TITLE                             = "About Donation Ledger Spreadsheet";

// Payment types
// P1 Regular tax deductible one time gift from PayPal
// P2 Monthly regular tax deductible from PayPal
// P3 Tax deductible with Comment from PayPal
// C1 regular check tax deductible
// D1 Donor Advised Fund check or EFT
// I1 IRA Distribution check
const PAYMENT_TYPE_P1 = "P1";
const PAYMENT_TYPE_P2 = "P2";
const PAYMENT_TYPE_P3 = "P3";
const PAYMENT_TYPE_C1 = "C1";
const PAYMENT_TYPE_D1 = "D1";
const PAYMENT_TYPE_I1 = "I1";

// Payment sources
const PAYMENT_SOURCE_PAYPAL ="PayPal";

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();

  ui
    .createMenu(NEWTON_TREE_CONSERVANCY_MENU)
      .addItem(IMPORT_PENDING_DONATION_DATA_MENU_ITEM, "onImportPendingDonationData")
      .addSeparator()
      .addItem(ABOUT_MENU_ITEM, "onAbout")
      .addToUi();
}

function onScheduledImport() {
  onImportPendingDonationData(false);
}

function onImportPendingDonationData(displayResult = true) {
  let sheet          = getDonationDataSheet_();
  let pendingFolder  = null;
  let importedFolder = null;
  let firstDataRow   = -1;

  let pendingFolderRange = sheet.getRange(PENDING_FOLDER_RANGE);

  if (pendingFolderRange != undefined) {
    try {
      pendingFolder = DriveApp.getFolderById(pendingFolderRange.getValue());
    }
    catch (e) {
      console.log("Pending folder is not accessible");
    }
  }
  else {
    console.log("Pending folder range is not defined");
  }

  let importedFolderRange = sheet.getRange(IMPORTED_FOLDER_RANGE);

  if (importedFolderRange != undefined) {
    try {
      importedFolder = DriveApp.getFolderById(importedFolderRange.getValue());
    }
    catch (e) {
      console.log("Imported folder is not accessible");
    }
  }
  else {
    console.log("Imported folder range is not defined");
  }

  let firstDataRowRange = sheet.getRange(FIRST_DATA_ROW_RANGE);

  if (firstDataRowRange != undefined) {
    firstDataRow = firstDataRowRange.getValue();
  }
  else {
    console.log("First data row range is not defined");
  }  

  if ((pendingFolder != null) && (importedFolder != null) && (firstDataRow != -1)) {
    let stats = importPendingFiles_(sheet, pendingFolder, importedFolder, firstDataRow, 1);

    let result = "";

    if (stats.length > 0) {
      result = "The following files (record counts) were imported:";

      result += "<ul>"

      stats.forEach(function(s) {
        result += `<li><p style="font-family:arial">${s[1]} (${s[2]})</p>`;
      });

      result += "</ul>"
    }
    else {
      result = "No files were pending import";
    }

    if (displayResult) {
      let ui   = SpreadsheetApp.getUi();
      let html = HtmlService.createHtmlOutput(`<p style="font-family:arial">${result}</p>`);
      
      ui.showModelessDialog(html, IMPORT_PENDING_DONATION_DATA_TITLE);
    }
  }
}

function onAbout() {
  let ui = SpreadsheetApp.getUi();

  ui.alert(ABOUT_TITLE,
    `Deployment ID
    ${DEPLOYMENT_ID}
    
    Version
    ${DEPLOYMENT_VERSION}


    Newton Tree Conservancy
    www.newtontreeconservancy.org`,
    ui.ButtonSet.OK);
}

function importPendingFiles_(sheet, pendingFolder, importedFolder, firstDataRow, firstDataColumn) {
  let stats = [];

  let pendingFiles = sortPendingFiles_(pendingFolder.getFiles());

  pendingFiles.forEach(function(f) {
    let fileName = f.getName().toLowerCase();

    if (fileName.startsWith(PAYPAL_FILE_PREFIX)) {
      let stat = importPayPalCsv_(sheet, f, firstDataRow, firstDataColumn);

      if (stat[0] == true) {
        f.moveTo(importedFolder);
      }

      stats.push(stat);
    }
  });

  return stats;
}

function importPayPalCsv_(sheet, file, firstDataRow, firstDataColumn) {
  let aOkay = true;

  let numRows    = 0;
  let numColumns = 0;

  let data = Utilities.parseCsv(file.getBlob().getDataAsString());

  if (data[0].length == PAYPAL_FILE_FIELD_COUNT) {
    data.splice(0, 1);

    let rows = normalizePaypalData_(data);

    numRows    = rows.length;
    numColumns = rows[0].length;

    sheet.insertRows(firstDataRow, numRows);
    sheet.getRange(firstDataRow, firstDataColumn, numRows, numColumns).
      setValues(rows).
      setBackground(DONATION_NEW_ROW_BACKGROUND_COLOR);
  }
  else {
    aOkay = false;

    console.log(`File ${file.getName()} contains insufficient number of fields: expected ${PAYPAL_FILE_FIELD_COUNT} found ${data[0].length}.`);
  }

  return [aOkay, file.getName(), numRows];
}

function normalizePaypalData_(data) {
  //out:Admin|Donation date|Last name|First name|Salutation/Other Names|Gross|Fee|Net|Payment type|Payment source |Payment Notes|Email address|Street address|City|State|Zip code

  let rows = [];

  data.forEach(function(r) {
    let donationType = r[4];

    if (isPayPalDonation(donationType)) {
      let row = [];

      // Admin
      row.push("");

      // Donation date
      row.push(r[0]);

      // Shipping address tokenizes the first and last "name" fields, so it's the preferred source for their values
      let shippingAddress = r[13].trim();
      let lastName        = "";
      let firstName       = "";

      if (shippingAddress.length > 0) {
        let tokens = shippingAddress.split(/\s*,\s*/)

        lastName  = tokens[1];
        firstName = tokens[0];
      }
      else if (donationType == PAYPAL_MASS_PAYMENT) {
        lastName  = r[3].trim();
        firstName = "";
      }
      else {
        let tokens = r[3].trim().split(/\s+/);

        if (tokens.length > 1) {
          lastName  = tokens[tokens.length - 1].trim();
          firstName = tokens.slice(0, -1).join(" ");
        }
        else {
          lastName  = tokens[0].trim();
          firstName = "";
        }
      }

      if (lastName.search(/\s+/) == -1) {
        lastName = lastName.toLocaleLowerCase();
      }

      if (firstName.search(/\s+/) == -1) {
        firstName = firstName.toLocaleLowerCase();
      }

      // Last name
      row.push(lastName.charAt(0).toUpperCase() + lastName.slice(1));

      // First name
      row.push(firstName.charAt(0).toUpperCase() + firstName.slice(1));

      // Salutation/Other names
      row.push("");

      // Gross
      row.push(r[7]);
      
      // Fee
      row.push(r[8]);
      
      // Net
      row.push(r[9]);
      
      // Payment type
      let paymentType = PAYMENT_TYPE_P1;

      if (donationType == PAYPAL_DONATION_PAYMENT) {
        if (r[38].trim().length > 0) {
          paymentType = PAYMENT_TYPE_P3;
        }
      }
      else if (donationType == PAYPAL_SUBSCRIPTION_PAYMENT) {
        paymentType = PAYMENT_TYPE_P2;
      }
      else if (donationType == PAYPAL_MASS_PAYMENT) {
        paymentType = PAYMENT_TYPE_D1;
      }

      row.push(paymentType);
      
      // Payment source 
      row.push(PAYMENT_SOURCE_PAYPAL);
      
      // Payment Notes
      row.push(r[38].trim());
      
      // Email address
      row.push(r[10].trim());

      let streetAddress1 = r[30].trim();
      let streetAddress2 = r[31].trim();

      if (streetAddress2.length > 0) {
        streetAddress1 += ADDRESS_JOIN_SEPARATOR + streetAddress2;
      } 

      // Street address
      row.push(streetAddress1);
      
      // City
      row.push(r[32].trim());
      
      // State
      row.push(r[33].trim());
      
      // Zip code
      row.push(r[34].trim());

      rows.push(row);
    }
  });

  if (rows.length > 1) {
    rows.sort((a, b) => new Date(b[1]) - new Date(a[1]));
  }

  return rows;
}

function isPayPalDonation(donationType) {
  return ((donationType == PAYPAL_DONATION_PAYMENT) || 
          (donationType == PAYPAL_SUBSCRIPTION_PAYMENT) ||
          (donationType == PAYPAL_MOBILE_PAYMENT) ||
          (donationType == PAYPAL_MASS_PAYMENT))
}

function sortPendingFiles_(unsortedFiles) {
  let sortedFiles = [];

  while (unsortedFiles.hasNext()) {
    sortedFiles.push(unsortedFiles.next());
  }

  if (sortedFiles.length > 1) {
    sortedFiles.sort((a, b) => a.getName().toLowerCase().localeCompare(b.getName().toLowerCase()));
  }

  return sortedFiles;
}

function getDonationDataSheet_() {
  let sheet = undefined;
  let range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(DONATION_DATA_RANGE);

  if (range != null) {
    sheet = range.getSheet();
  }
  else {
    throw new Error("Donation data sheet could not be resolved");
  }

  return sheet;
}
