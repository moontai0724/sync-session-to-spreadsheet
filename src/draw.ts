export function drawSolidBorder(
  range: GoogleAppsScript.Spreadsheet.Range,
  color = "black",
) {
  return range.setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    color,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
  );
}
