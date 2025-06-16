/* exported BANDING_THEME, getTemplate, getSubject */

const CONFIG_SHEET_NAME = "FRee Mail";

const CELL = {
  email: {
    range: "A2:A",
  },
  header: {
    all: {
      range: "1:1",
    },
    email: {
      note: "The email to send the draft to",
      range: "A1",
      value: "Email",
    },
    placeholder: {
      note: "The placeholder to replace with the values below",
      range: "B1:Z1",
      value: "{YOUR_PLACEHOLDER}",
    },
    subject: {
      note: "The subject of the emails to send",
      range: `${CONFIG_SHEET_NAME}!B1`,
      value: "Subject",
    },
    templateSubject: {
      note: "The subject of the template draft email",
      range: `${CONFIG_SHEET_NAME}!A1`,
      value: "Template Subject",
    },
  },
  placeholder: {
    range: "B2:Z",
  },
  subject: {
    range: `${CONFIG_SHEET_NAME}!B2`,
  },
  templateSubject: {
    range: `${CONFIG_SHEET_NAME}!A2`,
  },
};

const BANDING_THEME = SpreadsheetApp.BandingTheme.GREEN;

const getTemplate = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): string => {
  const subject = ss.getRange(CELL.templateSubject.range).getValue() as unknown;

  if (typeof subject !== "string") {
    throw new Error("Failed to get template subject.");
  }

  const threads = GmailApp.search(`subject:"${subject}"`);

  const body = threads[0]?.getMessages()[0]?.getBody();

  if (typeof body !== "string") {
    throw new Error(`No mail with subject "${subject}" found!`);
  }

  return body;
};

const getSubject = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): string => {
  const subject = ss.getRange(CELL.subject.range).getValue() as unknown;

  if (typeof subject !== "string") {
    throw new Error("Failed to get subject.");
  }

  return subject;
};
