/************************************************************
 * JYOTI GAS ENQUIRY SYSTEM
 *
 * GMAIL → PENDING → CONTACTED
 *
 * IMPORT START DATE:
 * 27 AUGUST 2026
 *
 * SHEETS:
 * Pending
 * Contacted
 *
 * COLUMNS:
 * A = Date
 * B = Name
 * C = Email
 * D = Phone
 * E = Subject
 * F = Message
 * G = Original Email Subject
 * H = Location
 * I = Status
 ************************************************************/


/************************************************************
 * SETTINGS
 ************************************************************/

const PENDING_SHEET = "Pending";
const CONTACTED_SHEET = "Contacted";

/*
 * Gmail will import emails from:
 *
 * 27 August 2026 onward
 *
 * Do NOT change this unless you want a different
 * starting date.
 */
const GMAIL_SEARCH =
  'subject:"Jyoti Gas" "New Enquiry" after:2026/08/26';


/************************************************************
 * 1. IMPORT NEW ENQUIRIES FROM GMAIL
 ************************************************************/

function importNewEnquiries() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const pendingSheet =
    ss.getSheetByName(PENDING_SHEET);

  if (!pendingSheet) {
    throw new Error(
      "Sheet named 'Pending' was not found."
    );
  }


  /*
   * Search Gmail.
   *
   * Only emails from 27 August 2026 onward
   * will be considered.
   */

  const threads =
    GmailApp.search(GMAIL_SEARCH);


  /*
   * Used to remember emails that have already
   * been imported.
   */

  const properties =
    PropertiesService.getScriptProperties();


  let addedCount = 0;


  /**********************************************************
   * LOOP THROUGH GMAIL THREADS
   **********************************************************/

  for (
    let t = 0;
    t < threads.length;
    t++
  ) {

    const messages =
      threads[t].getMessages();


    /********************************************************
     * LOOP THROUGH EMAILS
     ********************************************************/

    for (
      let m = 0;
      m < messages.length;
      m++
    ) {

      const message =
        messages[m];


      const messageId =
        message.getId();


      /******************************************************
       * DON'T IMPORT THE SAME EMAIL TWICE
       ******************************************************/

      if (
        properties.getProperty(
          "processed_" + messageId
        )
      ) {

        continue;
      }


      /******************************************************
       * GET EMAIL SUBJECT
       ******************************************************/

      const emailSubject =
        message.getSubject();


      const lowerSubject =
        emailSubject.toLowerCase();


      /******************************************************
       * ONLY PROCESS JYOTI GAS NEW ENQUIRIES
       ******************************************************/

      if (
        lowerSubject.indexOf("jyoti gas") === -1 ||
        lowerSubject.indexOf("new enquiry") === -1
      ) {

        continue;
      }


      /******************************************************
       * CHECK EMAIL DATE
       *
       * Extra safety check.
       *
       * Anything before 27 Aug 2026 is ignored.
       ******************************************************/

      const emailDate =
        message.getDate();


      const startDate =
        new Date("2026-08-27T00:00:00");


      if (
        emailDate < startDate
      ) {

        continue;
      }


      /******************************************************
       * GET EMAIL BODY
       ******************************************************/

      const body =
        message.getPlainBody();


      /******************************************************
       * EXTRACT INFORMATION
       ******************************************************/

      let name =
        getField(
          body,
          [
            "Name",
            "Customer Name"
          ]
        );


      let email =
        getField(
          body,
          [
            "Email",
            "Email ID",
            "Email Address"
          ]
        );


      let phone =
        getField(
          body,
          [
            "Phone",
            "Mobile",
            "Mobile Number",
            "Phone Number"
          ]
        );


      let subject =
        getField(
          body,
          [
            "Subject",
            "Enquiry Subject"
          ]
        );


      let enquiryMessage =
        getField(
          body,
          [
            "Message",
            "Enquiry Message"
          ]
        );


      let location =
        getField(
          body,
          [
            "Location",
            "City"
          ]
        );


      /******************************************************
       * DEFAULT VALUES
       ******************************************************/

      if (!name) {
        name = "No";
      }


      if (!email) {
        email = "No";
      }


      if (!phone) {
        phone = "No";
      }


      if (!subject) {
        subject = "General";
      }


      /*
       * MESSAGE RULE:
       *
       * No message = No
       *
       * Message exists = actual message
       */

      if (!enquiryMessage) {
        enquiryMessage = "No";
      }


      if (!location) {
        location = "No";
      }


      /******************************************************
       * CREATE ROW
       ******************************************************/

      const rowData = [

        emailDate,          // A - Date
        name,               // B - Name
        email,              // C - Email
        phone,              // D - Phone
        subject,            // E - Subject
        enquiryMessage,     // F - Message
        emailSubject,       // G - Original Email Subject
        location,           // H - Location
        "Pending"           // I - Status

      ];


      /******************************************************
       * ADD TO PENDING SHEET
       ******************************************************/

      pendingSheet.appendRow(
        rowData
      );


      /******************************************************
       * REMEMBER THIS EMAIL
       ******************************************************/

      properties.setProperty(
        "processed_" + messageId,
        "true"
      );


      /*
       * Mark Gmail email as read.
       *
       * If you want the emails to stay unread,
       * delete/comment the next line.
       */

      message.markRead();


      addedCount++;

    }
  }


  Logger.log(
    "Import completed."
  );

  Logger.log(
    "New enquiries added: " +
    addedCount
  );
}


/************************************************************
 * 2. EXTRACT FIELD FROM EMAIL
 *
 * Supports:
 *
 * Name: John
 * Name : John
 * Name - John
 * Name = John
 *
 ************************************************************/

function getField(
  body,
  fieldNames
) {

  if (!body) {
    return "";
  }


  const lines =
    body.split(/\r?\n/);


  for (
    let i = 0;
    i < lines.length;
    i++
  ) {

    const line =
      lines[i].trim();


    if (!line) {
      continue;
    }


    for (
      let f = 0;
      f < fieldNames.length;
      f++
    ) {

      const fieldName =
        fieldNames[f];


      const regex =
        new RegExp(
          "^" +
          escapeRegex(fieldName) +
          "\\s*[:\\-=]\\s*(.*)$",
          "i"
        );


      const match =
        line.match(regex);


      if (match) {

        return match[1].trim();
      }
    }
  }


  return "";
}


/************************************************************
 * 3. ESCAPE REGEX CHARACTERS
 ************************************************************/

function escapeRegex(text) {

  return text.replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&"
  );
}


/************************************************************
 * 4. MOVE CONTACTED ROW
 *
 * When Column I is changed to "Contacted",
 * the entire row moves from Pending to Contacted.
 ************************************************************/

function onEdit(e) {

  if (
    !e ||
    !e.range
  ) {

    return;
  }


  const range =
    e.range;


  const sheet =
    range.getSheet();


  /******************************************************
   * ONLY COLUMN I
   ******************************************************/

  if (
    range.getColumn() !== 9
  ) {

    return;
  }


  /******************************************************
   * DON'T PROCESS HEADER
   ******************************************************/

  if (
    range.getRow() === 1
  ) {

    return;
  }


  /******************************************************
   * DON'T PROCESS CONTACTED SHEET
   ******************************************************/

  if (
    sheet.getName() === CONTACTED_SHEET
  ) {

    return;
  }


  /******************************************************
   * GET STATUS
   ******************************************************/

  const status =
    range
      .getValue()
      .toString()
      .trim()
      .toLowerCase();


  /******************************************************
   * ONLY MOVE WHEN STATUS = CONTACTED
   ******************************************************/

  if (
    status !== "contacted"
  ) {

    return;
  }


  const ss =
    e.source;


  /******************************************************
   * FIND CONTACTED SHEET
   ******************************************************/

  let contactedSheet =
    ss.getSheetByName(
      CONTACTED_SHEET
    );


  /******************************************************
   * CREATE CONTACTED SHEET IF MISSING
   ******************************************************/

  if (!contactedSheet) {

    contactedSheet =
      ss.insertSheet(
        CONTACTED_SHEET
      );
  }


  const lastColumn =
    sheet.getLastColumn();


  /******************************************************
   * COPY HEADER
   ******************************************************/

  if (
    contactedSheet.getLastRow() === 0
  ) {

    const header =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getValues();


    contactedSheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .setValues(header);
  }


  /******************************************************
   * GET ROW
   ******************************************************/

  const rowNumber =
    range.getRow();


  const rowData =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastColumn
      )
      .getValues();


  /******************************************************
   * ADD TO CONTACTED
   ******************************************************/

  contactedSheet
    .getRange(
      contactedSheet.getLastRow() + 1,
      1,
      1,
      lastColumn
    )
    .setValues(
      rowData
    );


  /******************************************************
   * DELETE FROM PENDING
   ******************************************************/

  sheet.deleteRow(
    rowNumber
  );
}


/************************************************************
 * 5. MOVE EXISTING CONTACTED ROWS
 *
 * This is optional.
 *
 * Use this if Pending already has rows whose status
 * is Contacted.
 ************************************************************/

function moveExistingContacted() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const pendingSheet =
    ss.getSheetByName(
      PENDING_SHEET
    );


  if (!pendingSheet) {

    throw new Error(
      "Pending sheet was not found."
    );
  }


  let contactedSheet =
    ss.getSheetByName(
      CONTACTED_SHEET
    );


  if (!contactedSheet) {

    contactedSheet =
      ss.insertSheet(
        CONTACTED_SHEET
      );
  }


  const lastRow =
    pendingSheet.getLastRow();


  const lastColumn =
    pendingSheet.getLastColumn();


  /******************************************************
   * COPY HEADER
   ******************************************************/

  if (
    contactedSheet.getLastRow() === 0
  ) {

    const header =
      pendingSheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getValues();


    contactedSheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .setValues(
        header
      );
  }


  /******************************************************
   * MOVE CONTACTED ROWS
   *
   * Bottom to top prevents skipped rows.
   ******************************************************/

  for (
    let row = lastRow;
    row >= 2;
    row--
  ) {

    const status =
      pendingSheet
        .getRange(
          row,
          9
        )
        .getValue()
        .toString()
        .trim()
        .toLowerCase();


    if (
      status === "contacted"
    ) {

      const rowData =
        pendingSheet
          .getRange(
            row,
            1,
            1,
            lastColumn
          )
          .getValues();


      contactedSheet
        .getRange(
          contactedSheet.getLastRow() + 1,
          1,
          1,
          lastColumn
        )
        .setValues(
          rowData
        );


      pendingSheet.deleteRow(
        row
      );
    }
  }


  Logger.log(
    "Existing Contacted rows moved."
  );
}


/************************************************************
 * 6. TEST IMPORT
 *
 * Run this manually for testing.
 ************************************************************/

function testImport() {

  importNewEnquiries();

  Logger.log(
    "Test completed."
  );
}


/************************************************************
 * 7. CREATE AUTOMATIC 5-MINUTE TRIGGER
 *
 * Run this ONE TIME only.
 ************************************************************/

function createEmailTrigger() {

  const triggers =
    ScriptApp.getProjectTriggers();


  /******************************************************
   * REMOVE OLD IMPORT TRIGGERS
   ******************************************************/

  for (
    let i = 0;
    i < triggers.length;
    i++
  ) {

    if (
      triggers[i].getHandlerFunction() ===
      "importNewEnquiries"
    ) {

      ScriptApp.deleteTrigger(
        triggers[i]
      );
    }
  }


  /******************************************************
   * CREATE NEW 5-MINUTE TRIGGER
   ******************************************************/

  ScriptApp
    .newTrigger(
      "importNewEnquiries"
    )
    .timeBased()
    .everyMinutes(5)
    .create();


  Logger.log(
    "Automatic Gmail trigger created."
  );
}


/************************************************************
 * 8. SYSTEM TEST
 ************************************************************/

function systemTest() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const pending =
    ss.getSheetByName(
      PENDING_SHEET
    );


  const contacted =
    ss.getSheetByName(
      CONTACTED_SHEET
    );


  Logger.log(
    "Pending sheet: " +
    (
      pending
        ? "FOUND"
        : "NOT FOUND"
    )
  );


  Logger.log(
    "Contacted sheet: " +
    (
      contacted
        ? "FOUND"
        : "NOT FOUND"
    )
  );


  const threads =
    GmailApp.search(
      GMAIL_SEARCH
    );


  Logger.log(
    "Matching Gmail threads: " +
    threads.length
  );
}
