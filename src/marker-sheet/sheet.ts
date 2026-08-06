export const SESSION_PRIORITIES = [
  "important",
  "necessary",
  "notable",
] as const;

export type SessionPriority = typeof SESSION_PRIORITIES[number];

const SHEET_NAME = "Session Markers";
const HEADER_ROW = 1;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const MIN_DATA_ROWS = 100;
const INPUT_COLUMN_COUNT = 2;
const DETAIL_START_COLUMN = 3;
const DETAIL_COLUMN_COUNT = 3;
const COLUMN_COUNT = INPUT_COLUMN_COUNT + DETAIL_COLUMN_COUNT;

const PRIORITY_COLORS: Record<SessionPriority, string> = {
  important: "#F4CCCC",
  necessary: "#FCE5CD",
  notable: "#CFE2F3",
};

export class MarkerSheetManager {
  public readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly sessionsById: Map<EventSessionId, EventSession>;
  private readonly roomsById: Map<EventRoomId, EventRoom>;

  public constructor({ sessions, rooms }: EventData) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = this.spreadsheet.getSheetByName(SHEET_NAME);
    this.sheet = existingSheet ?? this.spreadsheet.insertSheet(SHEET_NAME);
    if (!existingSheet) resetSheet(this.sheet);

    this.sessionsById = new Map(sessions.map(session => [session.id, session]));
    this.roomsById = new Map(rooms.map(room => [room.id, room]));
  }

  public initialize(): void {
    this.ensureSize();
    this.renderHeaders();
    this.formatColumns();
    this.applyPriorityValidation();
    this.applyConditionalFormatting();
    this.fillSessionDetails();
    this.sheet.setFrozenRows(HEADER_ROW);
  }

  public getPriorities(): Map<EventSessionId, SessionPriority> {
    const priorities = new Map<EventSessionId, SessionPriority>();
    const seenSessionIds = new Set<EventSessionId>();
    const rowCount = this.sheet.getLastRow() - HEADER_ROW;
    if (rowCount <= 0) return priorities;

    const rows = this.sheet
      .getRange(FIRST_DATA_ROW, 1, rowCount, INPUT_COLUMN_COUNT)
      .getDisplayValues();
    rows.forEach(([rawSessionId, rawPriority], offset) => {
      const sessionId = rawSessionId.trim();
      const priority = rawPriority.trim();
      if (!sessionId && !priority) return;
      const row = FIRST_DATA_ROW + offset;
      if (!sessionId) throw new RangeError(`Missing session ID at row ${row}`);
      if (seenSessionIds.has(sessionId)) {
        throw new RangeError(
          `Duplicate session ID at row ${row}: ${sessionId}`,
        );
      }
      seenSessionIds.add(sessionId);
      if (!priority) return;
      if (!isSessionPriority(priority)) {
        throw new RangeError(`Invalid priority at row ${row}: ${priority}`);
      }
      priorities.set(sessionId, priority);
    });

    return priorities;
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
      .requireValueInList(Array.from(SESSION_PRIORITIES), true)
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
    const rules = SESSION_PRIORITIES.map(priority =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$B${FIRST_DATA_ROW}="${priority}"`)
        .setBackground(PRIORITY_COLORS[priority])
        .setRanges([range])
        .build(),
    );
    this.sheet.setConditionalFormatRules(rules);
  }

  private fillSessionDetails(): void {
    const rowCount = this.sheet.getLastRow() - HEADER_ROW;
    if (rowCount <= 0) return;

    const timezone = this.spreadsheet.getSpreadsheetTimeZone();
    const sessionIds = this.sheet
      .getRange(FIRST_DATA_ROW, 1, rowCount, 1)
      .getDisplayValues();
    const details = sessionIds.map(([rawSessionId]) => {
      const sessionId = rawSessionId.trim();
      if (!sessionId) return ["", "", ""];

      const session = this.sessionsById.get(sessionId);
      if (!session) {
        Logger.log(`WARN: Unknown marker session ID ${sessionId}.`);
        return ["", "", ""];
      }

      const room = this.roomsById.get(session.room);
      if (!room) Logger.log(`WARN: Session ${sessionId} has unknown room.`);

      return [
        session.zh.title,
        room?.zh.name ?? session.room,
        formatSessionTime(session, timezone),
      ];
    });

    this.sheet
      .getRange(
        FIRST_DATA_ROW,
        DETAIL_START_COLUMN,
        rowCount,
        DETAIL_COLUMN_COUNT,
      )
      .setValues(details);
  }
}

function isSessionPriority(value: string): value is SessionPriority {
  return (SESSION_PRIORITIES as readonly string[]).includes(value);
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

function formatSessionTime(session: EventSession, timezone: string): string {
  const startsAt = new Date(session.start);
  const endsAt = new Date(session.end);
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
