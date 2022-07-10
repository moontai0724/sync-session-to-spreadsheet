export default class SessionSheetManager {
  public readonly TIME_COLUMN = 1;
  public readonly ROOM_ROW = 3;
  public readonly UNIT_TIME_MINUTE = 5;
  public spreadsheet;
  public sheet;
  public roomColumnReferance: Record<EventRoomId, number>;
  public spacingColumns: number[];
  public baseTime: Date;

  public constructor(
    public sheetName: string,
    public date: string,
    public startHour: number,
    public endHour: number,
    public data: EventData,
    public importantSessions: EventSession[],
  ) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = this.spreadsheet.getSheetByName(this.sheetName);
    this.sheet = sheet ?? this.createSheet();
    this.roomColumnReferance = this.getRoomColumnReferance();
    this.spacingColumns = this.getSpacingColumns();
    this.startHour = parseInt(
      this.sheet
        .getRange(this.ROOM_ROW + 1, this.TIME_COLUMN)
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

  public createSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const sheet = this.spreadsheet.insertSheet(this.sheetName);

    // Reset sheet
    sheet.deleteColumns(1, sheet.getMaxColumns() - 1);
    sheet.deleteRows(1, sheet.getMaxRows() - 1);

    // Set rooms
    sheet.insertRowsAfter(1, this.ROOM_ROW - 1);
    const roomIds = this.data.rooms.map(room => room.id);
    sheet.insertColumnsAfter(1, roomIds.length);
    sheet
      .getRange(this.ROOM_ROW, 1, 1, roomIds.length + 1)
      .setValues([["時\\地", ...roomIds]]);

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
        const minute = (i * this.UNIT_TIME_MINUTE).toString().padStart(2, "0");
        const time = `${hour}:${minute}`;
        sheet.getRange(rowIndex, this.TIME_COLUMN).setValue(time);
      }
      const hourRange = sheet.getRange(
        baseRow,
        this.TIME_COLUMN,
        rowAmountOfHour,
        1,
      );
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
    sheet.setColumnWidth(this.TIME_COLUMN, 50);
    sheet.setFrozenRows(this.ROOM_ROW);
    sheet.setFrozenColumns(this.TIME_COLUMN);
    sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    return sheet;
  }

  public getRoomColumnReferance(): Record<EventRoomId, number> {
    const roomIds = this.data.rooms.map(room => room.id);
    const roomColumnReferance = roomIds.reduce((all, roomId) => {
      const matchCell = this.sheet
        .getRange(this.ROOM_ROW, 1, 1, this.sheet.getMaxColumns())
        .createTextFinder(roomId)
        .matchEntireCell(true)
        .findNext();
      if (!matchCell) return all;

      const columnIndex = matchCell.getColumn();
      return { ...all, [roomId]: columnIndex };
    }, {});

    return roomColumnReferance;
  }

  public getSpacingColumns(): number[] {
    const spacingColumns = this.sheet
      .getRange(this.ROOM_ROW, 1, 1, this.sheet.getMaxColumns())
      .createTextFinder("拍攝者")
      .findAll()
      .map(cell => cell.getColumn());

    return spacingColumns;
  }

  public clearCurrentSessions(): void {
    const sessionColumns = Object.values(this.roomColumnReferance);
    for (const column of sessionColumns) {
      const maxRow = this.sheet.getMaxRows();
      this.sheet
        .getRange(this.ROOM_ROW + 1, column, maxRow - this.ROOM_ROW, 1)
        .clear();
    }
  }

  public fillData(): void {
    this.clearCurrentSessions();
    this.data.sessions.forEach(session => {
      const start = new Date(session.start);

      if (start.toLocaleDateString() !== this.baseTime.toLocaleDateString())
        return;

      const column = this.roomColumnReferance[session.room];
      const startRow = this.getRowIndexOfTime(session.start);
      const endRow = this.getRowIndexOfTime(session.end) - 1;

      const range = this.sheet.getRange(
        startRow,
        column,
        endRow - startRow + 1,
        1,
      );

      const richValue = SpreadsheetApp.newRichTextValue()
        .setText(session.zh.title)
        .setLinkUrl(session.uri)
        .build();

      range
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
    });
    this.hightlightSessions();
    this.normalizeBorder();
  }

  public getRowIndexOfTime(time: string): number {
    const sessionTime = new Date(time);
    const offsetTime = sessionTime.getTime() - this.baseTime.getTime();
    const offsetRow = Math.round(
      offsetTime / 1000 / 60 / this.UNIT_TIME_MINUTE,
    );

    Logger.log(
      "Session time %s, base time %s, offset time %s, offset row %s",
      sessionTime.toLocaleString(),
      this.baseTime.toLocaleString(),
      offsetTime / 1000 / 60,
      offsetRow,
    );
    return this.ROOM_ROW + 1 + offsetRow;
  }

  public hightlightSessions(): void {
    this.importantSessions.forEach(session => {
      const start = new Date(session.start);
      if (start.toLocaleDateString() !== this.baseTime.toLocaleDateString())
        return;

      const column = this.roomColumnReferance[session.room];
      const startRow = this.getRowIndexOfTime(session.start);
      const endRow = this.getRowIndexOfTime(session.end) - 1;

      const range = this.sheet.getRange(
        startRow,
        column,
        endRow - startRow + 1,
        1,
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

  public normalizeBorder(): void {
    this.spacingColumns.forEach(column => {
      const maxRow = this.sheet.getMaxRows();
      this.sheet
        .getRange(this.ROOM_ROW + 1, column, maxRow - this.ROOM_ROW, 1)
        .setBorder(
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
