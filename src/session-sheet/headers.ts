export interface HeaderManagerOptions {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  rooms: readonly EventRoom[];
  fromRow: number;
  fromColumn: number;
  columnWidth?: number;
}

export class HeaderManager {
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public readonly rooms: readonly EventRoom[];
  public readonly fromRow: number;
  public readonly fromColumn: number;
  public readonly columnWidth: number;

  public constructor({
    sheet,
    rooms,
    fromRow,
    fromColumn,
    columnWidth = 150,
  }: HeaderManagerOptions) {
    if (!Number.isInteger(columnWidth) || columnWidth <= 0) {
      throw new RangeError("columnWidth must be a positive integer");
    }

    this.sheet = sheet;
    this.rooms = rooms;
    this.fromRow = fromRow;
    this.fromColumn = fromColumn;
    this.columnWidth = columnWidth;
  }

  public render(): void {
    if (this.rooms.length === 0) return;

    const lastColumn = this.fromColumn + this.rooms.length - 1;
    const missingColumns = lastColumn - this.sheet.getMaxColumns();
    if (missingColumns > 0) {
      this.sheet.insertColumnsAfter(this.sheet.getMaxColumns(), missingColumns);
    }

    const missingRows = this.fromRow - this.sheet.getMaxRows();
    if (missingRows > 0) {
      this.sheet.insertRowsAfter(this.sheet.getMaxRows(), missingRows);
    }

    this.sheet.setColumnWidths(
      this.fromColumn,
      this.rooms.length,
      this.columnWidth,
    );
    this.sheet
      .getRange(this.fromRow, this.fromColumn, 1, this.rooms.length)
      .setValues([this.rooms.map(room => room.zh.name)]);
  }

  public getColumnByRoomId(roomId: EventRoomId): number {
    const offset = this.rooms.findIndex(room => room.id === roomId);
    if (offset < 0) throw new RangeError(`Room ${roomId} is not rendered`);
    return this.fromColumn + offset;
  }
}
