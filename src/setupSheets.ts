/* exported insertConfigSheet, insertSampleSheet */

const applyFormat = (
  sheet: Readonly<GoogleAppsScript.Spreadsheet.Sheet>
): void => {
  const maxRange = sheet.getRange(
    1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  );

  sheet
    .getRange(CELL.header.all.range)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidths(1, maxRange.getLastColumn(), 200);
  maxRange.applyRowBanding(BANDING_THEME);
  sheet.setFrozenRows(1);
};

const insertConfigSheet = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): void => {
  const configSheet = ss.insertSheet(CONFIG_SHEET_NAME, 0);
  configSheet.deleteColumns(3, 24);
  configSheet.deleteRows(3, 998);

  configSheet
    .getRange(CELL.header.templateSubject.range)
    .setValue(CELL.header.templateSubject.value)
    .setNote(CELL.header.templateSubject.note);

  configSheet
    .getRange(CELL.header.subject.range)
    .setValue(CELL.header.subject.value)
    .setNote(CELL.header.subject.note);

  applyFormat(configSheet);
};

const insertSampleSheet = (
  ss: Readonly<GoogleAppsScript.Spreadsheet.Spreadsheet>
): void => {
  const sampleSheet = ss.insertSheet("Sample Sheet", 1);

  sampleSheet
    .getRange(CELL.header.email.range)
    .setValue(CELL.header.email.value)
    .setNote(CELL.header.email.note);

  sampleSheet
    .getRange(CELL.header.placeholder.range)
    .setValue(CELL.header.placeholder.value)
    .setNote(CELL.header.placeholder.note);

  applyFormat(sampleSheet);
};
