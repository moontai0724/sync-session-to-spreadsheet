import type { DailySessionManager } from "../data-manager/daily";
import { TimeManager } from "./time-slot";

const TIME_HEADER_ROW = 3;
const MINUTES_PER_TIME_SLOT = 5;

export class SessionSheetManager {
  public readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public readonly sheetName: string;
  public readonly timeManager: TimeManager;

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

    this.timeManager = new TimeManager({
      sheet: this.sheet,
      fromRow: TIME_HEADER_ROW,
      startAt: this.dailySessionManager.startsAt,
      endAt: this.dailySessionManager.endsAt,
      minutesPerUnit: MINUTES_PER_TIME_SLOT,
    });

    if (!existingSheet) {
      resetSheet(this.sheet);
      this.timeManager.render();
    }
  }
}

function resetSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.clear();
  sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
  sheet.deleteRows(2, sheet.getMaxRows() - 1);
}
