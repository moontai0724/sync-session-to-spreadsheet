import type { EventDayManager } from "../data-manager/event-day";
import type { Session } from "../data-manager/session";
import type { SessionManager } from "../data-manager/sessions";

export const SESSION_PRIORITIES = [
  "important",
  "necessary",
  "notable",
] as const;
export const SESSION_MARKER_LEVELS = [
  "hidden",
  "special",
  ...SESSION_PRIORITIES,
] as const;

export type SessionPriority = typeof SESSION_PRIORITIES[number];
export type SessionMarkerLevel = typeof SESSION_MARKER_LEVELS[number];

export interface SessionMarker {
  readonly priority?: SessionPriority;
  readonly hidden: boolean;
  readonly special: boolean;
}

const SHEET_NAME = "Session Markers";
const HEADER_ROW = 1;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const MIN_DATA_ROWS = 100;
const INPUT_COLUMN_COUNT = 2;
const DETAIL_START_COLUMN = 3;
const DETAIL_COLUMN_COUNT = 3;
const COLUMN_COUNT = INPUT_COLUMN_COUNT + DETAIL_COLUMN_COUNT;

export const SESSION_PRIORITY_COLORS: Record<SessionPriority, string> = {
  important: "#F4CCCC",
  necessary: "#FCE5CD",
  notable: "#CFE2F3",
};
export const SPECIAL_MARKER_BACKGROUND_COLOR = "#B6D7A8";
export const SPECIAL_SESSION_BORDER_COLOR = "#CC0000";

export class MarkerSheetManager {
  public readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  public constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = this.spreadsheet.getSheetByName(SHEET_NAME);
    this.sheet = existingSheet ?? this.spreadsheet.insertSheet(SHEET_NAME);
    if (!existingSheet) resetSheet(this.sheet);
  }

  public initialize(): void {
    this.ensureSize();
    this.renderHeaders();
    this.formatColumns();
    this.applyPriorityValidation();
    this.applyConditionalFormatting();
    this.sheet.setFrozenRows(HEADER_ROW);
  }

  public syncSessionDetails(
    sessionManager: SessionManager,
    eventDayManager: EventDayManager,
  ): void {
    const rowCount = this.sheet.getLastRow() - HEADER_ROW;
    if (rowCount <= 0) return;

    const timezone = this.spreadsheet.getSpreadsheetTimeZone();
    const sessionIds = this.sheet
      .getRange(FIRST_DATA_ROW, 1, rowCount, 1)
      .getDisplayValues();
    const details = sessionIds.map(([rawSessionId]) => {
      const sessionId = rawSessionId.trim();
      if (!sessionId) return { title: "", url: "", room: "", time: "" };

      const session = sessionManager.getSession(sessionId);
      if (!session) {
        Logger.log(`WARN: Unknown marker session ID ${sessionId}.`);
        return { title: "", url: "", room: "", time: "" };
      }

      return {
        title: session.title,
        url: eventDayManager.getSessionUrl(session.id),
        room: session.room?.zh.name ?? session.roomId,
        time: formatSessionTime(session, timezone),
      };
    });

    this.sheet
      .getRange(FIRST_DATA_ROW, DETAIL_START_COLUMN, rowCount, 1)
      .setRichTextValues(
        details.map(({ title, url }) => [createTitleValue(title, url)]),
      );
    this.sheet
      .getRange(
        FIRST_DATA_ROW,
        DETAIL_START_COLUMN + 1,
        rowCount,
        DETAIL_COLUMN_COUNT - 1,
      )
      .setValues(details.map(({ room, time }) => [room, time]));
  }

  public getMarkers(): Map<EventSessionId, SessionMarker> {
    const markers = new Map<EventSessionId, SessionMarker>();
    const rowCount = this.sheet.getLastRow() - HEADER_ROW;
    if (rowCount <= 0) return markers;

    const rows = this.sheet
      .getRange(FIRST_DATA_ROW, 1, rowCount, INPUT_COLUMN_COUNT)
      .getDisplayValues();
    rows.forEach(([rawSessionId, rawLevel], offset) => {
      const sessionId = rawSessionId.trim();
      const level = rawLevel.trim();
      if (!sessionId && !level) return;
      const row = FIRST_DATA_ROW + offset;
      if (!sessionId) throw new RangeError(`Missing session ID at row ${row}`);
      if (!level) return;
      if (!isSessionMarkerLevel(level)) {
        throw new RangeError(`Invalid marker level at row ${row}: ${level}`);
      }

      const current = markers.get(sessionId);
      if (level === "hidden") {
        markers.set(sessionId, {
          ...current,
          hidden: true,
          special: current?.special ?? false,
        });
        return;
      }

      if (level === "special") {
        markers.set(sessionId, {
          ...current,
          hidden: current?.hidden ?? false,
          special: true,
        });
        return;
      }

      markers.set(sessionId, {
        priority: getHigherPriority(current?.priority, level),
        hidden: current?.hidden ?? false,
        special: current?.special ?? false,
      });
    });

    return markers;
  }

  private ensureSize(): void {
    const requiredRows = HEADER_ROW + MIN_DATA_ROWS;
    const missingRows = requiredRows - this.sheet.getMaxRows();
    if (missingRows > 0) {
      this.sheet.insertRowsAfter(this.sheet.getMaxRows(), missingRows);
    }

    const missingColumns = COLUMN_COUNT - this.sheet.getMaxColumns();
    if (missingColumns > 0) {
      this.sheet.insertColumnsAfter(this.sheet.getMaxColumns(), missingColumns);
    }
  }

  private renderHeaders(): void {
    this.sheet
      .getRange(HEADER_ROW, 1, 1, COLUMN_COUNT)
      .setValues([["Session ID", "Priority", "Title", "Room", "Time"]])
      .setFontWeight("bold")
      .setBackground("#D9EAD3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  private formatColumns(): void {
    this.sheet.setColumnWidth(1, 160);
    this.sheet.setColumnWidth(2, 130);
    this.sheet.setColumnWidth(3, 300);
    this.sheet.setColumnWidth(4, 160);
    this.sheet.setColumnWidth(5, 220);
    this.sheet
      .getRange(FIRST_DATA_ROW, 1, this.sheet.getMaxRows() - HEADER_ROW, 1)
      .setNumberFormat("@");
    this.sheet
      .getRange(
        FIRST_DATA_ROW,
        1,
        this.sheet.getMaxRows() - HEADER_ROW,
        COLUMN_COUNT,
      )
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
    this.sheet
      .getRange(
        FIRST_DATA_ROW,
        DETAIL_START_COLUMN,
        this.sheet.getMaxRows() - HEADER_ROW,
        1,
      )
      .setHorizontalAlignment("left");
  }

  private applyPriorityValidation(): void {
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(Array.from(SESSION_MARKER_LEVELS), true)
      .setAllowInvalid(false)
      .build();
    this.sheet
      .getRange(FIRST_DATA_ROW, 2, this.sheet.getMaxRows() - HEADER_ROW, 1)
      .setDataValidation(validation);
  }

  private applyConditionalFormatting(): void {
    const range = this.sheet.getRange(
      FIRST_DATA_ROW,
      1,
      this.sheet.getMaxRows() - HEADER_ROW,
      COLUMN_COUNT,
    );
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$B${FIRST_DATA_ROW}="special"`)
        .setBackground(SPECIAL_MARKER_BACKGROUND_COLOR)
        .setRanges([range])
        .build(),
      ...SESSION_PRIORITIES.map(priority =>
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=$B${FIRST_DATA_ROW}="${priority}"`)
          .setBackground(SESSION_PRIORITY_COLORS[priority])
          .setRanges([range])
          .build(),
      ),
    ];
    this.sheet.setConditionalFormatRules(rules);
  }
}

function isSessionMarkerLevel(value: string): value is SessionMarkerLevel {
  return (SESSION_MARKER_LEVELS as readonly string[]).includes(value);
}

function getHigherPriority(
  current: SessionPriority | undefined,
  candidate: SessionPriority,
): SessionPriority {
  if (!current) return candidate;
  return SESSION_PRIORITIES.indexOf(candidate) <
    SESSION_PRIORITIES.indexOf(current)
    ? candidate
    : current;
}

function resetSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.clear();

  const requiredRows = HEADER_ROW + MIN_DATA_ROWS;
  const extraRows = sheet.getMaxRows() - requiredRows;
  if (extraRows > 0) sheet.deleteRows(requiredRows + 1, extraRows);

  const extraColumns = sheet.getMaxColumns() - COLUMN_COUNT;
  if (extraColumns > 0) {
    sheet.deleteColumns(COLUMN_COUNT + 1, extraColumns);
  }
}

function createTitleValue(
  title: string,
  url: string,
): GoogleAppsScript.Spreadsheet.RichTextValue {
  const builder = SpreadsheetApp.newRichTextValue().setText(title);
  if (url) builder.setLinkUrl(url);
  return builder.build();
}

function formatSessionTime(session: Session, timezone: string): string {
  const { startsAt, endsAt } = session;
  if (Number.isNaN(startsAt.getTime()) || Number.isNaN(endsAt.getTime())) {
    Logger.log(`WARN: Session ${session.id} has an invalid time.`);
    return "";
  }

  const startDate = Utilities.formatDate(startsAt, timezone, "yyyy-MM-dd");
  const endDate = Utilities.formatDate(endsAt, timezone, "yyyy-MM-dd");
  const startTime = Utilities.formatDate(startsAt, timezone, "HH:mm");
  const endTime = Utilities.formatDate(endsAt, timezone, "HH:mm");
  if (startDate === endDate) return `${startDate} ${startTime}–${endTime}`;
  return `${startDate} ${startTime}–${endDate} ${endTime}`;
}
