//
// Copyright 2026 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This library includes functions to help manage processing of reservation requests for
// tree sapling giveaways.
//
// Note: The email body text references the date and information of a giveaway. Be sure to
// UPDATE RES_ACK_EVENT_DATE and RES_ACK_EVENT_INFO_LINK for each giveaway.
//
// @OnlyCurrentDoc
//
const RES_ACK_EVENT_DATE           = "October 24, 2026";
const RES_ACK_EVENT_INFO_LINK      = "https://mailchi.mp/094d1fab6b72/2026-nahanton-sapling-giveaway"; 
const RES_ACK_EMAIL_SENDER_NAME    = "Newton Tree Conservancy";
const RES_ACK_EMAIL_REPLY_TO       = "newtontreeconservancy@gmail.com";
const RES_ACK_EMAIL_SUBJECT        = "Reservation received for the Newton Tree Conservancy tree sapling giveaway";
const RES_ACK_EMAIL_BODY_TEMPLATE  =
  `<!DOCTYPE html>
  <html>
  <body>
  <p>
  Thank you for submitting a reservation to receive tree saplings during the Newton Tree Conservancy's giveaway at Nahanton Park's Community Gardens on <strong><?!=RES_ACK_EVENT_DATE?></strong>. Your reservation specified a <strong>time slot at <?!=timeSlot?></strong> to select your saplings. If you have to cancel or change the time of your reservation, please inform the Newton Tree Conservancy in a reply to this email.
  </p>
  <p>
  To view information about the giveaway, visit
  <a href="<?!=RES_ACK_EVENT_INFO_LINK?>"><?!=new Date().getFullYear()?> Tree Sapling Giveaway</a>.
  </p>
  <p>
  See you soon,
  </p>
  <p>
  Your friends at Newton Tree Conservancy
  <br><a href="https://www.newtontreeconservancy.org/">www.newtontreeconservancy.org</a>
  </p>
  </body>
  </html>`;

const EMAIL_ADDRESS_RANGE = "email_address";
const TIME_SLOT_RANGE     = "time_slot";

function onSubmit(e) {
  let sheet         = e.range.getSheet();
  let rowIndex      = e.range.getRow();
  let emailAddress  = sheet.getRange(rowIndex, sheet.getRange(EMAIL_ADDRESS_RANGE).getColumn()).getValue();
  let timeSlotRange = sheet.getRange(rowIndex, sheet.getRange(TIME_SLOT_RANGE).getColumn());
  let timeSlot      = timeSlotRange.getValue();

  if ((emailAddress != undefined) && (timeSlot != undefined)) {
    timeSlotRange.setHorizontalAlignment("right");

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