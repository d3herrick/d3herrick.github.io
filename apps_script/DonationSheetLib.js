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
// This library requires the following oauthScopes:
//
// "oauthScopes": ["https://www.googleapis.com/auth/spreadsheets.currentonly",
//                 "https://www.googleapis.com/auth/spreadsheets",
//                 "https://www.googleapis.com/auth/drive.readonly",
//                 "https://www.googleapis.com/auth/drive",
//                 "https://www.googleapis.com/auth/script.container.ui",
//                 "https://www.googleapis.com/auth/script.send_mail"]
//
const DEPLOYMENT_ID                         = "1cXoHvwTUh5pTV3_0YHl9jZsL4YZ7Ie6juG307YwOBxGLjeF81khFYHcy";
const DEPLOYMENT_VERSION                    = "3";
const DONATION_DATA_RANGE                   = "donation_data";
const PENDING_FOLDER_RANGE                  = "pending_folder";
const IMPORTED_FOLDER_RANGE                 = "imported_folder";
const ACK_FOLDER_RANGE                      = "acknowledgement_folder";
const ZIPCODE_MINIMUM_LENGTH                = 5;
const PROCESSING_EMAIL_DIST_LIST_RANGE      = "processing_email_dist_list"
const PROCESSING_EMAIL_SENDER_NAME_RANGE    = "processing_email_sender_name"
const PROCESSING_EMAIL_REPLY_TO_RANGE       = "processing_email_reply_to";
const PROCESSING_RESULT_BODY_TEMPLATE_RANGE = "processing_email_body_template";
const IMPORT_RESULT_EMAIL_SUBJECT_RANGE     = "import_result_email_subject";
const ACK_BODY_TEMPLATE_RANGE               = "acknowledgement_body_template";
const ACK_SUBJECT_RANGE                     = "acknowledgement_subject";
const ACK_RESULT_EMAIL_SUBJECT_RANGE        = "acknowledgement_result_email_subject";
const NTC_FIRST_DATA_ROW_RANGE              = "ntc_first_data_row";
const PAYPAL_FIRST_DATA_ROW_RANGE           = "paypal_first_data_row";
const ADDRESS_JOIN_SEPARATOR                = ", ";
const PAYPAL_FILE_PREFIX                    = "paypal";
const PAYPAL_FILE_FIELD_COUNT               = 41;
const PAYPAL_DONATION_PAYMENT               = "Donation Payment";
const PAYPAL_SUBSCRIPTION_PAYMENT           = "Subscription Payment";
const PAYPAL_MOBILE_PAYMENT                 = "Mobile Payment";
const PAYPAL_MASS_PAYMENT                   = "Mass Pay Payment";
const CHECK_FILE_PREFIX                     = "check";
const NEWTON_TREE_CONSERVANCY_MENU          = "Newton Tree Conservancy";
const ABOUT_MENU_ITEM                       = "About...";
const IMPORT_DONATION_DATA_MENU_ITEM        = "Import donation data";
const IMPORT_DONATION_DATA_TITLE            = "Import Donation Data";
const GENERATE_DONATION_ACKS_MENU_ITEM      = "Generate donation acknowledgements";
const GENERATE_DONATION_ACKS_DATA_TITLE     = "Generate Donation Acknowledgements";
const ABOUT_TITLE                           = "About Donation Ledger Spreadsheet";

// HTML inline styles. The modeless dialog widget seems unable to process style elements, thus these inline styles
const STYLE_STANDARD_FONT     = "font-family:arial;font-size:14px";
const STYLE_MONOSPACED_FONT   = "font-family:courier;font-size:14px";
const STYLE_NO_WRAP           = "white-space: nowrap";
const STYLE_TABLE             = "border-collapse: collapse;border: 2px solid;letter-spacing: 1px";
const STYLE_TABLE_HEADER_CELL = "vertical-align:bottom;background-color: LightGray;border: 1px solid;text-align:left;vertical-align:bottom";
const STYLE_TABLE_ROW_CELL    = "border: 1px solid;text-align:left;vertical-align:top";

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
const PAYMENT_SOURCE_PAYPAL = "PayPal";

function onOpen(e) {
  let ui = SpreadsheetApp.getUi();

  ui.
    createMenu(NEWTON_TREE_CONSERVANCY_MENU).
      addItem(IMPORT_DONATION_DATA_MENU_ITEM, "onImportDonationData").
      addItem(GENERATE_DONATION_ACKS_MENU_ITEM, "onGenerateDonationAcks").
      addSeparator().
      addItem(ABOUT_MENU_ITEM, "onAbout").
      addToUi();
}

function onScheduledImport() {
  onImportDonationData(false, true);
}

function onScheduledGenerateDonationAcks() {
  onGenerateDonationAcks(false, true);
}

function onImportDonationData(displayResult = true, emailResult = true) {
  let sheet          = getDonationDataSheet_();
  let pendingFolder  = null;
  let importedFolder = null;

  let pendingFolderRange = sheet.getRange(PENDING_FOLDER_RANGE);

  if (pendingFolderRange != undefined) {
    pendingFolder = DriveApp.getFolderById(pendingFolderRange.getValue());
  }
  else {
    throw Error(`${PENDING_FOLDER_RANGE} named range is not defined`);
  }

  let importedFolderRange = sheet.getRange(IMPORTED_FOLDER_RANGE);

  if (importedFolderRange != undefined) {
    importedFolder = DriveApp.getFolderById(importedFolderRange.getValue());
  }
  else {
    throw Error(`${IMPORTED_FOLDER_RANGE} named range is not defined`);
  }

  if ((pendingFolder != null) && (importedFolder != null)) {
    let stats  = importPendingFiles_(sheet, pendingFolder, importedFolder);
    let result = "";

    if (stats.length > 0) {
      result = `<p style="${STYLE_STANDARD_FONT}">The following files (total/payments) were imported:</p>`;

      result += "<ul>"

      stats.forEach(function(s) {
        result += `<li style="${STYLE_MONOSPACED_FONT}">${s.fileStats[1]} (${s.fileStats[2]}/${s.fileStats[3]})</li>`;
      });

      result += "</ul>"

      stats.forEach(function(s) {
        if (s.paymentNotes.length > 0) {
          result += `<p><p style="${STYLE_STANDARD_FONT}">The following donations included a payment note:</p></p>`;
          result += 
            `<p>
            <table style="${STYLE_STANDARD_FONT};${STYLE_TABLE}">
            <tr>
            <th style="${STYLE_TABLE_HEADER_CELL};${STYLE_NO_WRAP}">Donation date
            <th style="${STYLE_TABLE_HEADER_CELL};">Last name
            <th style="${STYLE_TABLE_HEADER_CELL};">First name
            <th style="${STYLE_TABLE_HEADER_CELL};">Payment note
            </tr>`;

          s.paymentNotes.forEach(function (n) {
            result +=
              `<tr>
              <td style="${STYLE_TABLE_ROW_CELL}">${n[0]}</td>
              <td style="${STYLE_TABLE_ROW_CELL}">${n[1]}</td>
              <td style="${STYLE_TABLE_ROW_CELL}">${n[2]}</td>
              <td style="${STYLE_TABLE_ROW_CELL}">${n[3]}</td>
              </tr>`;
          });

          result += 
            `</table>
            </p>`;
        }
      });

      if (emailResult) {
        sendEmailProcessingResult_(sheet, sheet.getRange(IMPORT_RESULT_EMAIL_SUBJECT_RANGE).getValue(), result);
      }
    }
    else {
      result = `<p style="${STYLE_STANDARD_FONT}">No files were pending import</p>`;
    }

    if (displayResult) {
      displayProcessingResult_(IMPORT_DONATION_DATA_TITLE, result);
    }
  }
}

function onGenerateDonationAcks(displayResult = true, emailResult = true) {
  let sheet     = getDonationDataSheet_();
  let ackFolder = null;

  let ackFolderRange = sheet.getRange(ACK_FOLDER_RANGE);

  if (ackFolderRange != undefined) {
    ackFolder = DriveApp.getFolderById(ackFolderRange.getValue());
  }
  else {
    throw Error(`${ACK_FOLDER_RANGE} named range is not defined`);
  }

  if (ackFolder != null) {
    let stats  = generateDonationAcks_(sheet, ackFolder)
    let result = "";

    if (stats.ackStats.length > 0) {
      result = `<p style="${STYLE_STANDARD_FONT}">The following counts were recorded while processing donation acknowledgements:</p>`;

      result += "<ul>"
  
      result += `<li style="${STYLE_MONOSPACED_FONT}">Total donations processed........: ${stats.ackStats[0]}</li>`;
      result += `<li style="${STYLE_MONOSPACED_FONT}">Email acknowledgements sent......: ${stats.ackStats[1]}</li>`;
      result += `<li style="${STYLE_MONOSPACED_FONT}">Document acknowledgements created: ${stats.ackStats[2]}</li>`;
      result += `<li style="${STYLE_MONOSPACED_FONT}">Recurring donations skipped......: ${stats.ackStats[3]}</li>`;

      result += "</ul>"

      if (emailResult) {
        sendEmailProcessingResult_(sheet, sheet.getRange(ACK_RESULT_EMAIL_SUBJECT_RANGE).getValue(), result);
      }
    }
    else {
      result = `<p style="${STYLE_STANDARD_FONT}">No donations were pending processing</p>`;
    }

    if (displayResult) {
      displayProcessingResult_(GENERATE_DONATION_ACKS_DATA_TITLE, result);
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

function importPendingFiles_(sheet, pendingFolder, importedFolder) {
  let ntcFirstDataRow    = undefined;
  let ntcFirstDataColumn = 1;

  let ntcFirstDataRowRange = sheet.getRange(NTC_FIRST_DATA_ROW_RANGE);

  if (ntcFirstDataRowRange != undefined) {
    ntcFirstDataRow = ntcFirstDataRowRange.getValue();
  }
  else {
    throw Error(`${NTC_FIRST_DATA_ROW_RANGE} named range is not defined`);
  }  

  let paypalFirstDataRow    = undefined;
  let paypalFirstDataColumn = 1;

  let paypalFirstDataRowRange = sheet.getRange(PAYPAL_FIRST_DATA_ROW_RANGE);

  if (paypalFirstDataRowRange != undefined) {
    paypalFirstDataRow = paypalFirstDataRowRange.getValue();
  }
  else {
    throw Error(`${PAYPAL_FIRST_DATA_ROW_RANGE} named range is not defined`);
  }

  let firstInsertionRow    = ntcFirstDataRow;
  let firstInsertionColumn = ntcFirstDataColumn;

  let stats = [];

  let pendingFiles = sortPendingFiles_(pendingFolder.getFiles());

  pendingFiles.forEach(function(f) {
    let fileName = f.getName().toLowerCase();
    let stat     = null;

    if (fileName.startsWith(PAYPAL_FILE_PREFIX)) {
      stat = importPayPalCsv_(sheet, f, paypalFirstDataRow, paypalFirstDataColumn, firstInsertionRow, firstInsertionColumn);
    }
    else if (fileName.startsWith(CHECK_FILE_PREFIX)) {
      stat = importSpreadsheet_(sheet, f, ntcFirstDataRow, ntcFirstDataColumn, firstInsertionRow, firstInsertionColumn);
    }
    else {
      stat = {
        fileStats:    [false, `${fileName}`, 0, `${f.getName()} is an unsupported file type`],
        paymentNotes: []
      };
    }

    if (stat.fileStats[0] == true) {
      f.moveTo(importedFolder);
    }

    stats.push(stat);
  });

  return stats;
}

function importPayPalCsv_(sheet, file, firstDataRow, firstDataColumn, firstInsertionRow, firstInsertionColumn) {
  let aOkay = true;

  let totalRows    = 0;
  let numRows      = 0;
  let numColumns   = 0;
  let paymentNotes = [];

  let data = Utilities.parseCsv(file.getBlob().getDataAsString());

  if (data[0].length == PAYPAL_FILE_FIELD_COUNT) {
    totalRows = data.length;

    let donations = normalizePaypalData_(data, firstDataRow);

    numRows    = donations.length;
    numColumns = donations[0].length;

    insertDonationData_(sheet, donations, firstInsertionRow, firstInsertionColumn, numRows, numColumns);

    paymentNotes = tabulatePaymentNotes_(donations);
  }
  else {
    aOkay = false;

    numRows = `${file.getName()} contains unexpected number of fields: expected ${PAYPAL_FILE_FIELD_COUNT} found ${data[0].length}`;
  }

  return {
    fileStats:    [aOkay, file.getName(), totalRows, numRows],
    paymentNotes: paymentNotes
  };
}

function normalizePaypalData_(data, firstDataRow) {
  //out:Ack emailed or generated|Donation date|Last name|First name|Salutation/Other names|Gross|Fee|Net|Payment type|Payment source |Payment note|Email address|Street address|City|State|Zip code

  let rows = [];

  data.forEach(function(r) {
    if (!isEmptyRow_(r)) {
      let donationType = r[4];

      if (isPayPalDonation_(donationType)) {
        let row = [];

        // Ack emailed or generated
        row.push("");

        // Donation date
        row.push(normalizeDonationDate_(r[0]));

        // Shipping address tokenizes the its fields, including first and last "name", so it's the preferred source for those values
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
        
        // Payment note
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
        row.push(normalizeDonationZipcode_(r[34].trim()));

        rows.push(row);
      }
    }
  });

  return rows;
}

function importSpreadsheet_(sheet, file, firstDataRow, firstDataColumn, firstInsertionRow, firstInsertionColumn) {
  let aOkay = true;

  let totalRows    = 0;
  let numRows      = 0;
  let numColumns   = 0;
  let paymentNotes = [];

  let range = SpreadsheetApp.openById(file.getId()).getRangeByName(DONATION_DATA_RANGE);

  if (range != null) {
    range     = range.getSheet().getDataRange();
    totalRows = range.getNumRows();

    let donations = normalizeCheckData_(range.getValues(), firstDataRow);

    numRows    = donations.length;
    numColumns = donations[0].length;

    insertDonationData_(sheet, donations, firstInsertionRow, firstInsertionColumn, numRows, numColumns);

    paymentNotes = tabulatePaymentNotes_(donations);
  }
  else {
    aOkay = false;

    numRows = `${DONATION_DATA_RANGE} named range is not defined in ${file.getName()}`;
  }

  return {
    fileStats:    [aOkay, file.getName(), totalRows, numRows],
    paymentNotes: paymentNotes
  };
}

function normalizeCheckData_(data, firstDataRow) {
  //out:Ack emailed or generated|Donation date|Last name|First name|Salutation/Other names|Gross|Fee|Net|Payment type|Payment source |Payment note|Email address|Street address|City|State|Zip code

  data.splice(0, (firstDataRow - 1));

  let rows = [];

  data.forEach(function(r) {
    if (!isEmptyRow_(r)) {
      let row = [];

      // Ack emailed or generated
      row.push("");
      
      // Donation date
      row.push(normalizeDonationDate_(r[0]));

      // Last name
      row.push(r[1].toString().trim());

      // First name
      row.push(r[2].toString().trim());

      // Salutation/Other names
      row.push(r[3] = r[3].toString().trim());

      // Gross: copy net to gross
      row.push(r[6]);
      
      // Fee
      row.push(r[5]);
      
      // Net
      row.push(r[6]);

      // Payment type
      row.push(r[7].toString().trim());

      // Payment source
      row.push(r[8].toString().trim());
      
      // Payment note
      row.push(r[9].toString().trim());
      
      // Email address
      row.push(r[10].toString().trim());
      
      // Street address
      row.push(r[11].toString().trim());
      
      // City
      row.push(r[12].toString().trim());
      
      // State
      row.push(r[13].toString().trim());

      // Zip code
      row.push(normalizeDonationZipcode_(r[14]));

      rows.push(row);
    }
  });

  return rows;
}

function normalizeDonationDate_(date) {
  let normalizedDate = undefined;

  if (Number.isInteger(date)) {
    normalizedDate = date.toString();
    normalizedDate = `${normalizedDate.substring(0, 4)}/${normalizedDate.substring(4, 6)}/${normalizedDate.substring(6, 8)}`;
  }
  else {
    normalizedDate = date;
  }

  return normalizedDate;
}

function normalizeDonationZipcode_(zipCode) {
  let normalizedZipcode = undefined;

  if ((zipCode.length > 0) && (zipCode.length == (ZIPCODE_MINIMUM_LENGTH - 1))) {
    normalizedZipcode = "0" + zipCode;
  }
  else {
    normalizedZipcode = zipCode;
  }

  return normalizedZipcode;
}

function isEmptyRow_(row) {
  let isEmpty = true;

  if ((row != null) && (row.length > 0)) {
    for (let i = 0; i < row.length; i++) {
      if ((row[i] != null) && (row[i].toString().length > 0)) {
        isEmpty = false;
        break;
      }
    }
  }

  return isEmpty;
}

function tabulatePaymentNotes_(donations) {
  let paymentNotes = [];

  if (donations.length > 0) {
    donations.forEach(function (d) {
      let donationObject = toDonationObject_(d);

      if (donationObject.paymentNote.length > 0) {
        paymentNotes.push([donationObject.donationDate,
          donationObject.lastName,
          donationObject.firstName,
          donationObject.paymentNote
        ]);
      }
    });
  }

  return paymentNotes;
}

function generateDonationAcks_(sheet, ackFolder) {
  let ntcFirstDataRow    = undefined;
  let ntcFirstDataColumn = 1;
  
  let ntcFirstDataRowRange = sheet.getRange(NTC_FIRST_DATA_ROW_RANGE);

  if (ntcFirstDataRowRange != undefined) {
    ntcFirstDataRow = ntcFirstDataRowRange.getValue();
  }
  else {
    throw Error(`${NTC_FIRST_DATA_ROW_RANGE} named range is not defined`);
  }

  let spreadsheet   = sheet.getParent();
  let range         = sheet.getDataRange();
  let numRows       = range.getLastRow() - ntcFirstDataRow + 1;
  let numColumns    = range.getLastColumn();
  let ackRange      = sheet.getRange(ntcFirstDataRow, ntcFirstDataColumn, numRows, 1);
  let ackA1Notation = `'${sheet.getName()}'!${ackRange.getA1Notation()}`;

  let query = `=arrayformula(query({${ackA1Notation}, ROW(${ackA1Notation})}, "SELECT Col2 WHERE Col1 = FALSE", 0))`;

  let queryResults = spreadsheet.insertSheet().hideSheet();
  let ackData      = null;

  try {
    queryResults.getRange(1,1).setFormula(query);

    ackData = queryResults.getDataRange().getValues();
  }
  finally {
    spreadsheet.deleteSheet(queryResults);
  }

  let stats = [];

  if ((ackData.length > 0) && (ackData[0] != "#N/A")) {
    let emailProperties = {
      senderName: sheet.getRange(PROCESSING_EMAIL_SENDER_NAME_RANGE).getValue(),
      replyTo:    sheet.getRange(PROCESSING_EMAIL_REPLY_TO_RANGE).getValue(),
      subject:    sheet.getRange(ACK_SUBJECT_RANGE).getValue()
    };

    let bodyTemplate = HtmlService.createTemplateFromFile(sheet.getRange(ACK_BODY_TEMPLATE_RANGE).getValue());

    let totalAcks    = ackData.length;
    let numEmailAcks = 0;
    let numDocAcks   = 0;
    let numRecurring = 0;

    ackData.forEach(function (a) {
      let donationRange  = sheet.getRange(a[0], ntcFirstDataColumn, 1, numColumns);
      let donationObject = toDonationObject_(donationRange.getValues()[0]);

      bodyTemplate.donationDate    = donationObject.donationDate;
      bodyTemplate.lastName        = donationObject.lastName;
      bodyTemplate.firstName       = donationObject.firstName;
      bodyTemplate.salutation      = donationObject.salutation;
      bodyTemplate.paymentType     = donationObject.paymentType;
      bodyTemplate.paymentSource   = donationObject.paymentSource;
      bodyTemplate.emailAddress    = donationObject.emailAddress;
      bodyTemplate.streetAddress   = donationObject.streetAddress;
      bodyTemplate.city            = donationObject.city;
      bodyTemplate.state           = donationObject.state;
      bodyTemplate.zipCode         = donationObject.zipCode;

      let doGenerateAck = true;

      switch (donationObject.paymentType) {
        case PAYMENT_TYPE_P1:
        case PAYMENT_TYPE_P3:
          bodyTemplate.paymentAmount = donationObject.gross;
          break;

        case PAYMENT_TYPE_P2:
          doGenerateAck = false;
          numRecurring++;
          break;

        case PAYMENT_TYPE_C1:
        case PAYMENT_TYPE_D1:
        case PAYMENT_TYPE_I1:
          bodyTemplate.paymentAmount = donationObject.net;
          break;

        default:
          break;
      }

      if (doGenerateAck) {
        if (donationObject.emailAddress.length > 0) {
          sendEmailAck_(donationObject, bodyTemplate, emailProperties);
          numEmailAcks++;
        }
        else {
          createDocumentAck_(donationObject, bodyTemplate, ackFolder);
          numDocAcks++;
        }
      }

      sheet.getRange(a[0], ntcFirstDataColumn).check();
    });

    stats = [totalAcks, numEmailAcks, numDocAcks, numRecurring]; 
  }

  return {
    ackStats: stats
  };
}

function sendEmailAck_(donationObject, bodyTemplate, emailProperties) {
  bodyTemplate.doInlineImage = false;

  let body = bodyTemplate.evaluate().getContent();

  MailApp.sendEmail(
    donationObject.emailAddress,
    emailProperties.subject,
    null,
    {
      htmlBody: body,
      replyTo : emailProperties.replyTo,
      name    : emailProperties.senderName
    }
  );
}

function createDocumentAck_(donationObject, bodyTemplate, ackFolder) {
  bodyTemplate.doInlineImage = true;

  let body = bodyTemplate.evaluate().getContent();

  let docName = `${donationObject.lastName}_${donationObject.firstName}_${Intl.DateTimeFormat("en-US").format(donationObject.donationDate)}.pdf`;
  let docBlob = Utilities.newBlob(body, MimeType.HTML).getAs(MimeType.PDF).setName(docName);
  let docFile = DriveApp.createFile(docBlob);
  
  docFile.moveTo(ackFolder);
}

function displayProcessingResult_(subject, result) {
    let ui   = SpreadsheetApp.getUi();
    let html = HtmlService.createHtmlOutput(`<p style="${STYLE_STANDARD_FONT}">${result}</p>`);

    html.setHeight(400);
    html.setWidth(600);
    
    ui.showModelessDialog(html, subject);
}

function sendEmailProcessingResult_(sheet, subject, result) {
  let senderName       = sheet.getRange(PROCESSING_EMAIL_SENDER_NAME_RANGE).getValue();
  let distributionList = sheet.getRange(PROCESSING_EMAIL_DIST_LIST_RANGE).getValue();
  let replyTo          = sheet.getRange(PROCESSING_EMAIL_REPLY_TO_RANGE).getValue();
  let bodyTemplate     = HtmlService.createTemplateFromFile(sheet.getRange(PROCESSING_RESULT_BODY_TEMPLATE_RANGE).getValue());

  bodyTemplate.result = result +
    `<p><p style="${STYLE_STANDARD_FONT}">View the donation ledger by clicking <a href="${sheet.getParent().getUrl()}">here</a>.</p></p>`;

  let body = bodyTemplate.evaluate().getContent();

  MailApp.sendEmail(
    distributionList,
    subject,
    null,
    {
      htmlBody: body,
      replyTo : replyTo,
      name    : senderName
    }
  );
}

function toDonationObject_(donationRow) {
  let donationObject = {
    donationDate:  donationRow[1],
    lastName:      donationRow[2],
    firstName:     donationRow[3],
    salutation:    donationRow[4],
    gross:         donationRow[5],
    fee:           donationRow[6],
    net:           donationRow[7],
    paymentType:   donationRow[8],
    paymentSource: donationRow[9],
    paymentNote:   donationRow[10],
    emailAddress:  donationRow[11],
    streetAddress: donationRow[12],
    city:          donationRow[13],
    state:         donationRow[14],
    zipCode:       donationRow[15]
  };

  return donationObject;
}

function isPayPalDonation_(donationType) {
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

function insertDonationData_(sheet, donations, firstInsertionRow, firstInsertionColumn, numRows, numColumns) {
  if (donations.length > 0) {
    if (donations.length > 1) {
      donations.sort((a, b) => new Date(b[1]) - new Date(a[1]));
    }

    sheet.insertRows(firstInsertionRow, numRows);
    sheet.getRange(firstInsertionRow, firstInsertionColumn, numRows, numColumns).
      setValues(donations);
    sheet.getRange(firstInsertionRow, firstInsertionColumn, numRows, firstInsertionColumn).
      insertCheckboxes();
  }
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
