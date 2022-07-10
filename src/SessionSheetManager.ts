export default class SessionSheetManager {
  public readonly TIME_COLUMN = 1;
  public readonly ROOM_ROW = 3;
  public readonly UNIT_TIME_MINUTE = 5;
  public spreadsheet;
  public sheet;

  public constructor(
    public sheetName: string,
    public date: string,
    public startHour: number,
    public endHour: number,
    public data: EventData,
  ) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = this.spreadsheet.getSheetByName(this.sheetName);
    this.sheet = sheet ?? this.createSheet();
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
}
