import type { DailySessionManager } from "../data-manager/daily";

export class SessionSheetManager {
  public readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public readonly sheetName: string;

  public constructor(
    public readonly day: number,
    public readonly dailySessionManager: DailySessionManager,
  ) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const date = dailySessionManager.startsAt.toLocaleDateString("zh-TW", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });
    this.sheetName = `Day ${day} (${date})`;
    const existingSheet = this.spreadsheet.getSheetByName(this.sheetName);
    this.sheet = existingSheet ?? this.spreadsheet.insertSheet(this.sheetName);

    if (!existingSheet) {
      resetSheet(this.sheet);
    }
  }
}

function resetSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.clear();
  sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
  sheet.deleteRows(2, sheet.getMaxRows() - 1);
}
