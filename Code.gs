/************************************************************
 * TALLY → GOOGLE CALENDAR INVITATION TOOL
 *
 * HOW IT WORKS:
 * 1. Open the Tally response sheet you want to use.
 * 2. Go to:
 *      📅 Calendar Invitations
 *      → Create Calendar Invitations
 * 3. The script automatically uses the ACTIVE sheet.
 * 4. Configure the event in the modal.
 * 5. Preview the respondents.
 * 6. Send the invitations.
 *
 * TIMEZONE:
 * Asia/Manila
 ************************************************************/


const TIMEZONE = "Asia/Manila";


/************************************************************
 * DEFAULT EVENT SETTINGS
 *
 * These are the values that initially appear in the modal.
 *
 * You DO NOT need to edit these every time.
 * The user can change them directly in the modal.
 ************************************************************/

const DEFAULTS = {

  title:
    "🚀 Welcome to TechSTART: Think Tank AI!",

  date:
    "2026-10-21",

  startTime:
    "13:00",

  endTime:
    "15:50",

  description:
`🚀 Welcome to TechSTART: Think Tank AI!

Join us for an interactive workshop on AI-assisted ideation and learn how to turn a problem into a structured, practical solution concept using generative AI and prompt engineering.

📅 October 21, 2026 | 1:00–3:50 PM PHT
💻 Online via Zoom
📌 Guide (READ THIS): https://tinyurl.com/START-TechStartOct-GuideBook

Before the event:
1. Confirm your attendance through this calendar invite (Yes, No, Maybe). 
2. Read the guide.
3. Make your own copy of the Canva workbook attached in the guide.
4. Go to Slide 2 and draft your ideas as you listen. Rough notes are fine!
5. Prepare a laptop or desktop, sign in to Canva, and open an AI tool (ChatGPT, Gemini, or Claude) in another tab.
6. Think of one problem you care about.

🔗 Join Zoom: https://us06web.zoom.us/j/86469263838?pwd=Q51bUsDwYHbDWcQAYfIktV1VWsqwqP.1
Meeting ID: 864 6926 3838
Passcode: 402299
`

};


/************************************************************
 * ADD MENU TO GOOGLE SHEETS
 ************************************************************/

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("📅 Calendar Invitations")
    .addItem(
      "Create Calendar Invitations",
      "showCalendarModal"
    )
    .addToUi();

}


/************************************************************
 * OPEN THE MODAL
 ************************************************************/

function showCalendarModal() {

  const html =
    HtmlService
      .createHtmlOutputFromFile(
        "CalendarDialog"
      )
      .setWidth(720)
      .setHeight(820);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "📅 Calendar Invitation Tool"
    );

}


/************************************************************
 * GET INITIAL DATA
 *
 * IMPORTANT:
 * This gets the CURRENTLY ACTIVE SHEET.
 *
 * Therefore:
 *
 * Sheet A → open tool → uses Sheet A
 *
 * Sheet B → open tool → uses Sheet B
 *
 * Sheet C → open tool → uses Sheet C
 ************************************************************/

function getInitialData() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();


  if (!sheet) {

    throw new Error(
      "No active sheet was found."
    );

  }


  const lastColumn =
    sheet.getLastColumn();

  const lastRow =
    sheet.getLastRow();


  if (
    lastColumn === 0 ||
    lastRow === 0
  ) {

    throw new Error(
      "The active sheet is empty."
    );

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(String);


  const emailIndex =
    findColumn(
      headers,
      "Email Address"
    );


  if (emailIndex === -1) {

    throw new Error(
      'The active sheet does not contain an "Email Address" column.'
    );

  }


  /**********************************************************
   * GET AVAILABLE CALENDARS
   **********************************************************/

  const calendars =
    CalendarApp
      .getAllCalendars();


  const calendarList =
    calendars
      .filter(function(calendar) {

        try {

          return calendar
            .isMyPrimaryCalendar() ||
            calendar
              .getAccessLevel() !==
              CalendarApp
                .AccessLevel
                .READ;

        } catch (error) {

          return true;

        }

      })
      .map(function(calendar) {

        return {

          id:
            calendar.getId(),

          name:
            calendar.getName()

        };

      })
      .sort(function(a, b) {

        return a.name.localeCompare(
          b.name
        );

      });


  if (
    calendarList.length === 0
  ) {

    throw new Error(
      "No Google Calendars were found."
    );

  }


  return {

    sheetName:
      sheet.getName(),

    rowCount:
      Math.max(
        0,
        lastRow - 1
      ),

    emailColumn:
      headers[emailIndex],

    firstNameColumn:
      getColumnName(
        headers,
        "First Name"
      ),

    lastNameColumn:
      getColumnName(
        headers,
        "Last Name"
      ),

    calendars:
      calendarList,

    defaults:
      DEFAULTS

  };

}


/************************************************************
 * GET ACTIVE SHEET INFORMATION
 ************************************************************/

function getActiveSheetInfo() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();


  if (!sheet) {

    throw new Error(
      "No active sheet was found."
    );

  }


  const lastColumn =
    sheet.getLastColumn();

  const lastRow =
    sheet.getLastRow();


  if (
    lastColumn === 0 ||
    lastRow === 0
  ) {

    throw new Error(
      "The active sheet is empty."
    );

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(String);


  const emailIndex =
    findColumn(
      headers,
      "Email Address"
    );


  if (emailIndex === -1) {

    throw new Error(
      'The active sheet does not contain an "Email Address" column.'
    );

  }


  return {

    sheetName:
      sheet.getName(),

    totalRows:
      Math.max(
        0,
        lastRow - 1
      ),

    emailColumn:
      headers[emailIndex],

    firstNameColumn:
      getColumnName(
        headers,
        "First Name"
      ),

    lastNameColumn:
      getColumnName(
        headers,
        "Last Name"
      )

  };

}


/************************************************************
 * PREVIEW INVITATIONS
 ************************************************************/

function previewInvitations(settings) {

  validateSettings(
    settings
  );


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();


  /*
   * Make sure the sheet has not changed
   * unexpectedly.
   */

  if (
    sheet.getName() !==
    settings.sheetName
  ) {

    throw new Error(
      "The active sheet changed while the tool was open. " +
      "Please close the tool and open it again from the sheet you want to use."
    );

  }


  const data =
    sheet
      .getDataRange()
      .getValues();


  if (
    data.length < 2
  ) {

    throw new Error(
      "No submissions were found in the active sheet."
    );

  }


  const headers =
    data[0].map(String);


  const emailIndex =
    findColumn(
      headers,
      "Email Address"
    );


  if (
    emailIndex === -1
  ) {

    throw new Error(
      'Could not find "Email Address".'
    );

  }


  /*
   * Get or create Invitation Status.
   */

  const statusIndex =
    getOrCreateColumn(
      sheet,
      headers,
      "Invitation Status"
    );


  const uniqueEmails =
    new Set();


  let missingEmail = 0;

  let invalidEmail = 0;

  let duplicates = 0;

  let alreadyInvited = 0;


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const email =
      String(
        row[emailIndex] || ""
      )
        .trim()
        .toLowerCase();


    const status =
      String(
        row[statusIndex] || ""
      )
        .trim();


    /*
     * Missing email.
     */

    if (!email) {

      missingEmail++;

      continue;

    }


    /*
     * Invalid email.
     */

    if (
      !isValidEmail(email)
    ) {

      invalidEmail++;

      continue;

    }


    /*
     * Already invited.
     */

    if (
      status.toLowerCase() ===
      "invited"
    ) {

      alreadyInvited++;

      continue;

    }


    /*
     * Duplicate email.
     */

    if (
      uniqueEmails.has(email)
    ) {

      duplicates++;

      continue;

    }


    uniqueEmails.add(
      email
    );

  }


  return {

    totalSubmissions:
      data.length - 1,

    validEmails:
      uniqueEmails.size,

    missingEmail:
      missingEmail,

    invalidEmail:
      invalidEmail,

    duplicates:
      duplicates,

    alreadyInvited:
      alreadyInvited

  };

}


/************************************************************
 * SEND INVITATIONS
 ************************************************************/

function sendInvitations(settings) {

  validateSettings(
    settings
  );


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();


  /*
   * Confirm that we're still on
   * the same sheet.
   */

  if (
    sheet.getName() !==
    settings.sheetName
  ) {

    throw new Error(
      "The active sheet changed while the tool was open. " +
      "Please close the tool and open it again."
    );

  }


  const data =
    sheet
      .getDataRange()
      .getValues();


  if (
    data.length < 2
  ) {

    throw new Error(
      "No submissions were found."
    );

  }


  const headers =
    data[0].map(String);


  const emailIndex =
    findColumn(
      headers,
      "Email Address"
    );


  const firstNameIndex =
    findColumn(
      headers,
      "First Name"
    );


  const lastNameIndex =
    findColumn(
      headers,
      "Last Name"
    );


  if (
    emailIndex === -1
  ) {

    throw new Error(
      'Could not find "Email Address".'
    );

  }


  /*
   * Get or create Invitation Status.
   */

  const statusIndex =
    getOrCreateColumn(
      sheet,
      headers,
      "Invitation Status"
    );


  /*
   * Collect valid unique recipients.
   */

  const recipients = [];

  const seenEmails =
    new Set();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const email =
      String(
        row[emailIndex] || ""
      )
        .trim()
        .toLowerCase();


    const status =
      String(
        row[statusIndex] || ""
      )
        .trim();


    /*
     * Ignore empty emails.
     */

    if (!email) {

      continue;

    }


    /*
     * Ignore invalid emails.
     */

    if (
      !isValidEmail(email)
    ) {

      continue;

    }


    /*
     * Ignore people who have
     * already been marked as invited.
     */

    if (
      status.toLowerCase() ===
      "invited"
    ) {

      continue;

    }


    /*
     * Ignore duplicate emails.
     */

    if (
      seenEmails.has(email)
    ) {

      continue;

    }


    seenEmails.add(
      email
    );


    recipients.push({

      email:
        email,

      firstName:
        firstNameIndex !== -1
          ? String(
              row[firstNameIndex] || ""
            ).trim()
          : "",

      lastName:
        lastNameIndex !== -1
          ? String(
              row[lastNameIndex] || ""
            ).trim()
          : "",

      rowNumber:
        i + 1

    });

  }


  if (
    recipients.length === 0
  ) {

    throw new Error(
      "There are no new valid respondents to invite."
    );

  }


  /**********************************************************
   * TEST MODE
   **********************************************************/

  if (
    settings.testMode === true
  ) {

    return {

      testMode:
        true,

      eventCreated:
        false,

      recipientCount:
        recipients.length,

      message:
        "TEST MODE COMPLETE\n\n" +

        recipients.length +

        " valid unique respondents were found.\n\n" +

        "No calendar event was created.\n" +

        "No invitation emails were sent."

    };

  }


  /**********************************************************
   * CREATE DATE/TIME
   **********************************************************/

  const start =
    createDateTime(
      settings.date,
      settings.startTime
    );


  const end =
    createDateTime(
      settings.date,
      settings.endTime
    );


  if (
    end <= start
  ) {

    throw new Error(
      "End time must be later than start time."
    );

  }


  /**********************************************************
   * GET SELECTED CALENDAR
   **********************************************************/

  const calendar =
    CalendarApp
      .getCalendarById(
        settings.calendarId
      );


  if (!calendar) {

    throw new Error(
      "The selected calendar could not be found."
    );

  }


  /**********************************************************
   * CREATE ATTENDEE LIST
   **********************************************************/

  const attendeeEmails =
    recipients
      .map(function(person) {

        return person.email;

      });


  /*
   * Google Calendar accepts guests
   * as a comma-separated string.
   */

  const guestList =
    attendeeEmails.join(",");


  /**********************************************************
   * CREATE ONE CALENDAR EVENT
   **********************************************************/

  const event =
    calendar.createEvent(
      settings.title,
      start,
      end,
      {

        description:
          settings.description,

        guests:
          guestList,

        sendInvites:
          true

      }
    );


  /**********************************************************
   * MARK RESPONDENTS AS INVITED
   **********************************************************/

  let invitedCount = 0;

  let errorCount = 0;


  for (
    const person of recipients
  ) {

    try {

      sheet
        .getRange(
          person.rowNumber,
          statusIndex + 1
        )
        .setValue(
          "Invited"
        );

      invitedCount++;

    } catch (error) {

      sheet
        .getRange(
          person.rowNumber,
          statusIndex + 1
        )
        .setValue(
          "Error: " +
          error.message
        );

      errorCount++;

    }

  }


  /**********************************************************
   * RETURN RESULT
   **********************************************************/

  return {

    testMode:
      false,

    eventCreated:
      true,

    eventId:
      event.getId(),

    invitedCount:
      invitedCount,

    errorCount:
      errorCount,

    recipientCount:
      recipients.length,

    message:
      "Calendar event created successfully!\n\n" +

      "Event: " +
      settings.title +
      "\n\n" +

      "Date: " +
      formatDateForMessage(
        settings.date
      ) +
      "\n" +

      "Time: " +
      formatTime(
        settings.startTime
      ) +
      " – " +
      formatTime(
        settings.endTime
      ) +
      "\n\n" +

      "Invitation emails requested: " +
      invitedCount +
      "\n" +

      "Errors: " +
      errorCount

  };

}


/************************************************************
 * VALIDATE SETTINGS
 ************************************************************/

function validateSettings(settings) {

  if (!settings) {

    throw new Error(
      "No event settings were provided."
    );

  }


  if (!settings.sheetName) {

    throw new Error(
      "No active sheet was detected."
    );

  }


  if (!settings.calendarId) {

    throw new Error(
      "Please select a calendar."
    );

  }


  if (
    !settings.title ||
    settings.title.trim() === ""
  ) {

    throw new Error(
      "Please enter an event title."
    );

  }


  if (!settings.date) {

    throw new Error(
      "Please select an event date."
    );

  }


  if (
    !settings.startTime ||
    !settings.endTime
  ) {

    throw new Error(
      "Please enter both start and end times."
    );

  }

}


/************************************************************
 * FIND COLUMN
 ************************************************************/

function findColumn(
  headers,
  name
) {

  return headers.findIndex(
    function(header) {

      return header
        .trim()
        .toLowerCase() ===
        name
          .trim()
          .toLowerCase();

    }
  );

}


/************************************************************
 * GET COLUMN NAME
 ************************************************************/

function getColumnName(
  headers,
  name
) {

  const index =
    findColumn(
      headers,
      name
    );


  return index === -1
    ? ""
    : headers[index];

}


/************************************************************
 * GET OR CREATE COLUMN
 ************************************************************/

function getOrCreateColumn(
  sheet,
  headers,
  name
) {

  const existingIndex =
    findColumn(
      headers,
      name
    );


  if (
    existingIndex !== -1
  ) {

    return existingIndex;

  }


  const newColumn =
    sheet.getLastColumn() + 1;


  sheet
    .getRange(
      1,
      newColumn
    )
    .setValue(
      name
    );


  return newColumn - 1;

}


/************************************************************
 * VALIDATE EMAIL
 ************************************************************/

function isValidEmail(email) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(email);

}


/************************************************************
 * CREATE DATE/TIME
 ************************************************************/

function createDateTime(
  date,
  time
) {

  return Utilities.parseDate(
    `${date} ${time}:00`,
    TIMEZONE,
    "yyyy-MM-dd HH:mm:ss"
  );

}


/************************************************************
 * FORMAT TIME
 ************************************************************/

function formatTime(time) {

  const parts =
    time.split(":");


  let hour =
    parseInt(
      parts[0],
      10
    );


  const minute =
    parts[1];


  const period =
    hour >= 12
      ? "PM"
      : "AM";


  if (hour === 0) {

    hour = 12;

  } else if (hour > 12) {

    hour -= 12;

  }


  return `${hour}:${minute} ${period}`;

}


/************************************************************
 * FORMAT DATE
 ************************************************************/

function formatDateForMessage(
  dateString
) {

  const date =
    new Date(
      dateString +
      "T00:00:00"
    );


  return date.toLocaleDateString(
    "en-PH",
    {

      year:
        "numeric",

      month:
        "long",

      day:
        "numeric"

    }
  );

}
