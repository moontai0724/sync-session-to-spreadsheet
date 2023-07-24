export default class SessionSheetManager {
  /**
   * Column index of time, starts from 1.
   * Side effect: will read first row of time to be the base time to calculate where sessions are going to be put.
   * This only read hour of the time, this script is hour-based.
   */
  public readonly TIME_END_COLUMN = 2;
  /**
   * Row index of room, starts from 1.
   * Side effect: will default the line after `ROOM_ROW` is the first empty line for data to write.
   * Please leave rows below this row empty.
   */
  public readonly ROOM_ROW = 3;
  /** A unit time in minute, will used for init sheet and calculate where sessions are going to be put. */
  public readonly UNIT_TIME_MINUTE = 5;
  public spreadsheet;
  public sheet;
  public roomColumnReferance: Record<EventRoomId, number>;
  public spacingColumns: number[];
  public baseTime: Date;
  public roomTypes: Record<EventRoomId, EventSessionTypeId[]> = {};

  /**
   * @param sheetName Name of the sheet to be interact, can be a non-exist sheet, will auto create if so.
   * @param date Date of the event in this sheet, in format that parsable by `Date` object.
   * @param data Data of complete event.
   * @param importantSessions Important sessions that will be highlighted in red border and background.
   * @param startHour The first hour of this event, only used for create sheet, will be auto overwitten by the first row of time if the sheet already exists.
   * @param endHour The end hour of this sheet, is important when sheet is empty, no usage if sheet is not empty.
   */
  public constructor(
    public sheetName: string,
    public date: string,
    public data: EventData,
    public importantSessions: EventSession[] = [],
    public startHour: number = 0,
    public endHour: number = 24,
  ) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = this.spreadsheet.getSheetByName(this.sheetName);
    this.sheet = sheet ?? this.createSheet();
    this.roomColumnReferance = this.getRoomColumnReferance();
    this.spacingColumns = this.getSpacingColumns();
    this.startHour = parseInt(
      this.sheet
        .getRange(this.ROOM_ROW + 1, this.TIME_END_COLUMN - 1)
        .activate()
        .getDisplayValue()
        .split(":")[0],
      10,
    );
    this.baseTime = new Date(this.date + " " + this.startHour + ":00");
    Logger.log(
      "Sheet %s by referanced %s spacing by %s from %s to %s, base time is %s",
      this.sheetName,
      this.roomColumnReferance,
      this.spacingColumns,
      this.startHour,
      this.endHour,
      this.baseTime.toLocaleString(),
    );
  }

  /**
   * Create a default structure of sheet
   * includes rooms, times and styles.
   * @returns Sheet of created session
   */
  public createSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const sheet = this.spreadsheet.insertSheet(this.sheetName);

    // Reset sheet
    sheet.deleteColumns(1, sheet.getMaxColumns() - 1);
    sheet.deleteRows(1, sheet.getMaxRows() - 1);

    // Set rooms
    sheet.insertRowsAfter(1, this.ROOM_ROW - 1);
    const roomIds = this.data.rooms.map(room => room.id);
    sheet.insertColumnsAfter(1, this.TIME_END_COLUMN - 1 + roomIds.length);
    sheet
      .getRange(this.ROOM_ROW, this.TIME_END_COLUMN - 1, 1, roomIds.length + 2)
      .activate()
      .setValues([["開始", "結束", ...roomIds]]);

    // Set times
    const totalRows =
      ((this.endHour - this.startHour + 1) * 60) / this.UNIT_TIME_MINUTE;
    sheet.insertRowsAfter(this.ROOM_ROW, totalRows);

    for (let hour = this.startHour; hour <= this.endHour; hour++) {
      const rowAmountOfHour = 60 / this.UNIT_TIME_MINUTE;
      const baseRow =
        this.ROOM_ROW + 1 + (hour - this.startHour) * rowAmountOfHour;
      for (let i = 0; i < rowAmountOfHour; i++) {
        const rowIndex = baseRow + i;
        const startMinute = (i * this.UNIT_TIME_MINUTE)
          .toString()
          .padStart(2, "0");
        const startTime = `${hour}:${startMinute}`;
        const endMinute = (i * this.UNIT_TIME_MINUTE + this.UNIT_TIME_MINUTE)
          .toString()
          .padStart(2, "0");
        const endTime = `${hour}:${endMinute}`;
        sheet
          .getRange(rowIndex, this.TIME_END_COLUMN - 1)
          .setValue(startTime)
          .activate();
        sheet
          .getRange(rowIndex, this.TIME_END_COLUMN)
          .setValue(endTime)
          .activate();
      }
      const hourRange = sheet
        .getRange(baseRow, this.TIME_END_COLUMN - 1, rowAmountOfHour, 2)
        .activate();
      hourRange.setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        "black",
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
      );
    }

    // Beautify sheet
    sheet.setColumnWidth(this.TIME_END_COLUMN - 1, 50);
    sheet.setColumnWidth(this.TIME_END_COLUMN, 50);
    sheet.setFrozenRows(this.ROOM_ROW);
    sheet.setFrozenColumns(this.TIME_END_COLUMN);
    sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .activate()
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    return sheet;
  }

  /**
   * Summary and find referance between room and column index.
   * @returns A referance of room column, which is a map of room id to column index.
   */
  public getRoomColumnReferance(): Record<EventRoomId, number> {
    const roomIds = this.data.rooms.map(room => room.id);
    const roomColumnReferance = roomIds.reduce((all, roomId) => {
      const matchCell = this.sheet
        .getRange(this.ROOM_ROW, 1, 1, this.sheet.getMaxColumns())
        .activate()
        .createTextFinder(roomId)
        .matchEntireCell(true)
        .findNext();
      if (!matchCell) return all;

      const columnIndex = matchCell.getColumn();
      return { ...all, [roomId]: columnIndex };
    }, {});

    return roomColumnReferance;
  }

  /**
   * Get spacing columns.
   * If a column titled with keyword "拍攝者" that will be a spaceing column.
   * Will draw a black border for spacing columns later.
   * @returns An array of column index that is used for spacing.
   */
  public getSpacingColumns(): number[] {
    const spacingColumns = this.sheet
      .getRange(this.ROOM_ROW, 1, 1, this.sheet.getMaxColumns())
      .activate()
      .createTextFinder("拍攝者")
      .findAll()
      .map(cell => cell.getColumn());

    return spacingColumns;
  }

  /**
   * Clear current existing sessions in this sheet.
   * Only clear those sessions columns that are identified in RoomColumnReferance.
   */
  public clearCurrentSessions(): void {
    const sessionColumns = Object.values(this.roomColumnReferance);
    for (const column of sessionColumns) {
      const maxRow = this.sheet.getMaxRows();
      this.sheet
        .getRange(this.ROOM_ROW + 1, column, maxRow - this.ROOM_ROW, 1)
        .activate()
        .clear();
    }
  }

  /**
   * Fill sessions data into sheet.
   * Will cleare current sessions first.
   * Will fill sessions data into sheet.
   * Will draw black border for spacing columns.
   * Will hightlight sessions that are marked as important.
   */
  public fillData(): void {
    this.clearCurrentSessions();
    this.data.sessions.forEach(session => {
      const start = new Date(session.start);

      if (start.toLocaleDateString() !== this.baseTime.toLocaleDateString())
        return;

      const column = this.roomColumnReferance[session.room];
      if (!column) {
        Logger.log("[ERROR] No column found for room %s", session.room);
        return;
      }

      const startRow = this.getRowIndexOfTime(session.start);
      const endRow = this.getRowIndexOfTime(session.end) - 1;

      const range = this.sheet
        .getRange(startRow, column, endRow - startRow + 1, 1)
        .activate();
      Logger.log(
        "Fill session title=%s, url=%s, target=%s",
        session.zh.title,
        session.uri,
        range.getA1Notation(),
      );

      const richValue = SpreadsheetApp.newRichTextValue()
        .setText(session.zh.title)
        .setLinkUrl(session.uri)
        .build();

      // Following will cost too much time to execute in GAS
      // Only enable when there are conflicts and you want to preserve the first result.

      // const ranges = range.getMergedRanges();

      // if (ranges.length > 0) {
      //   Logger.log(
      //     "Error: Conflict! Target range %s have existing data!",
      //     range
      //       .getMergedRanges()
      //       .map(r => r.getA1Notation())
      //       .join(", "),
      //   );
      //   return;
      // }

      range
        .clear()
        .merge()
        .setRichTextValue(richValue)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true)
        .setBackground("#FDFDFD")
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          "black",
          SpreadsheetApp.BorderStyle.SOLID,
        );

      this.roomTypes[session.room] = [
        ...(this.roomTypes[session.room] ?? []),
        session.type,
      ];
    });
    this.setSessionType();
    this.hightlightSessions();
    this.normalizeBorder();
  }

  /**
   * Set a type for each session track (column) above row of room
   * by summary types of session in the track and find the most common type.
   */
  public setSessionType(): void {
    Logger.log("Start set session type");
    for (const roomId in this.roomTypes) {
      const types = this.roomTypes[roomId];
      const typeAmount = types.reduce(
        (all: Record<string, number>, value) => ({
          ...all,
          [value]: (all[value] ?? 0) + 1,
        }),
        {},
      );
      const maxAmount = Math.max(...Object.values(typeAmount));
      const typeId = Array.from(Object.entries(typeAmount)).find(
        current => current[1] === maxAmount,
      )?.[0];
      if (!typeId) continue;
      const type = this.data.session_types.find(type => type.id === typeId);
      if (!type) continue;
      const typeColumn = this.roomColumnReferance[roomId];

      const cellAboveRoom = this.sheet
        .getRange(this.ROOM_ROW - 1, typeColumn)
        .activate();

      if (cellAboveRoom.getDisplayValue() !== "") {
        Logger.log(
          "Warning: Cell above room %s is not empty to be filled with type, ignored to fill type.",
          roomId,
        );
        continue;
      }

      Logger.log(
        "Set type %s for room %s, column %s, target cell %s",
        type.zh.name,
        roomId,
        typeColumn,
        cellAboveRoom.getA1Notation(),
      );

      cellAboveRoom.setValue(type.zh.name);
    }
  }

  /**
   * Get a row index by time.
   * @param time A time string which parsable by Date object.
   * @returns Corresponding row index of the time.
   */
  public getRowIndexOfTime(time: string): number {
    const sessionTime = new Date(time);
    const offsetTime = sessionTime.getTime() - this.baseTime.getTime();
    const offsetRow = Math.round(
      offsetTime / 1000 / 60 / this.UNIT_TIME_MINUTE,
    );

    Logger.log(
      "Get session time %s, base time %s, offset time %s, offset row %s",
      sessionTime.toLocaleString(),
      this.baseTime.toLocaleString(),
      offsetTime / 1000 / 60,
      offsetRow,
    );
    return this.ROOM_ROW + 1 + offsetRow;
  }

  /**
   * Highlight sessions that are marked as important in red border and background.
   */
  public hightlightSessions(): void {
    Logger.log("Start highlight sessions");
    this.importantSessions.forEach(session => {
      const start = new Date(session.start);
      if (start.toLocaleDateString() !== this.baseTime.toLocaleDateString())
        return;

      const column = this.roomColumnReferance[session.room];
      const startRow = this.getRowIndexOfTime(session.start);
      const endRow = this.getRowIndexOfTime(session.end) - 1;

      const range = this.sheet
        .getRange(startRow, column, endRow - startRow + 1, 1)
        .activate();
      Logger.log(
        "Highlight session title=%s, url=%s, target=%s",
        session.zh.title,
        session.uri,
        range.getA1Notation(),
      );

      range
        .setBackground("#F4CCCC")
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          "red",
          SpreadsheetApp.BorderStyle.SOLID_THICK,
        );
    });
  }

  /**
   * Normalize border of spacing columns.
   * Will draw black, SOLID_THICK border for spacing columns.
   */
  public normalizeBorder(): void {
    Logger.log("Start normalize border");
    this.spacingColumns.forEach(column => {
      const maxRow = this.sheet.getMaxRows();
      const range = this.sheet
        .getRange(this.ROOM_ROW + 1, column, maxRow - this.ROOM_ROW, 1)
        .activate();
      Logger.log("Normalize border for range %s", range.getA1Notation());
      range.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "black",
        SpreadsheetApp.BorderStyle.SOLID_THICK,
      );
    });
  }
}
