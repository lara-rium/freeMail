/* exported createDrafts */

const createDrafts = (): void => {
  const ss = SpreadsheetApp.getActive();

  validate(ss);

  const template = getTemplate(ss);
  const subject = getSubject(ss);

  for (const sheet of ss.getSheets()) {
    if (sheet.getName() === CONFIG_SHEET_NAME) {
      continue;
    }

    processSheet(sheet, template, subject);
  }
};
