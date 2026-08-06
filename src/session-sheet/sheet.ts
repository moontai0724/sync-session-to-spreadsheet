import type { DailySessionManager } from "../data-manager/daily";
import type { SessionMarker } from "../marker-sheet";
import { HeaderManager } from "./headers";
import { SessionManager } from "./session";
import { TimeManager } from "./time-slot";

const HEADER_ROW = 3;
const MINUTES_PER_TIME_SLOT = 5;

export class SessionSheetManager {
  public readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public readonly sheetName: string;
  public readonly timeManager: TimeManager;
  public readonly headerManager: HeaderManager;
  public readonly sessionManager: SessionManager;

  public constructor(
    public readonly day: number,
    public readonly dailySessionManager: DailySessionManager,
    markers: ReadonlyMap<EventSessionId, SessionMarker>,
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
    if (!existingSheet) resetSheet(this.sheet);

    this.timeManager = new TimeManager({
      sheet: this.sheet,
      fromRow: HEADER_ROW,
      startAt: this.dailySessionManager.startsAt,
      endAt: this.dailySessionManager.endsAt,
      minutesPerUnit: MINUTES_PER_TIME_SLOT,
    });
    this.timeManager.render();

    this.headerManager = new HeaderManager({
      sheet: this.sheet,
      rooms: Array.from(this.dailySessionManager.activeRooms.values()),
      fromRow: HEADER_ROW,
      fromColumn: this.timeManager.fromColumn + 2,
    });
    this.headerManager.render();

    this.sessionManager = new SessionManager({
      sessions: this.dailySessionManager.sessions,
      timeManager: this.timeManager,
      headerManager: this.headerManager,
      markers,
    });
    this.sessionManager.render();

    if (!existingSheet) {
      this.sheet.setFrozenRows(this.headerManager.fromRow);
      this.sheet.setFrozenColumns(this.timeManager.fromColumn + 1);
    }
  }
}

function resetSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.clear();
  sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
  sheet.deleteRows(2, sheet.getMaxRows() - 1);
}
