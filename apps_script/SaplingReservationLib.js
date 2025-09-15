//
// Copyright 2025 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This library includes functions to help manage processing of reservation requests for tree sapling giveaways.
//
// Note: The email body text includes the specific dates of a giveaway. Be sure to UPDATE THE DATES FOR EACH GIVEAWAY.
//
// @OnlyCurrentDoc
//
const EMAIL_ADDRESS_RANGE         = "email_address";
const TIME_SLOT_RANGE             = "time_slot";
const RES_ACK_EMAIL_SENDER_NAME   = "Newton Tree Conservancy";
const RES_ACK_EMAIL_REPLY_TO      = "newtontreeconservancy@gmail.com";
const RES_ACK_EMAIL_SUBJECT       = "Reservation received for the Newton Tree Conservancy tree sapling giveaway";
const RES_ACK_EMAIL_BODY_TEMPLATE =
  `<!DOCTYPE html>
  <html>
  <body>
  <p>
  Thank you for submitting a reservation to receive tree saplings during the Newton Tree Conservancy's giveaway at Nahanton Park's Community Gardens on October 18, 2025. <strong>Your reservation specified a time slot at <?!=timeSlot?></strong> to receive your saplings.
  To view information about the giveaway, please visit
  <a href="https://www.newtontreeconservancy.org/news-content/2025-tree-sapling-giveaway">2025 Tree Sapling Giveaway</a>.
  </p>
  <p>
  </p>
  </body>
  </html>`;

function onSubmit(e) {
  let sheet        = e.range.getSheet();
  let rowIndex     = e.range.getRow();
  let emailAddress = sheet.getRange(rowIndex, sheet.getRange(EMAIL_ADDRESS_RANGE).getColumn()).getValue();
  let timeSlot     = sheet.getRange(rowIndex, sheet.getRange(TIME_SLOT_RANGE).getColumn()).getValue();

  if ((emailAddress != undefined) && (timeSlot != undefined)) {
    let senderName   = RES_ACK_EMAIL_SENDER_NAME;
    let replyTo      = RES_ACK_EMAIL_REPLY_TO;
    let subject      = RES_ACK_EMAIL_SUBJECT;
    let bodyTemplate = HtmlService.createTemplate(RES_ACK_EMAIL_BODY_TEMPLATE);
 
    bodyTemplate.timeSlot = timeSlot;

    let body = bodyTemplate.evaluate().getContent();

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