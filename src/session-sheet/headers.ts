export interface HeaderManagerOptions {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  rooms: readonly EventRoom[];
  fromRow: number;
  fromColumn: number;
}

export class HeaderManager {
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public readonly rooms: readonly EventRoom[];
  public readonly fromRow: number;
  public readonly fromColumn: number;

  public constructor({
    sheet,
    rooms,
    fromRow,
    fromColumn,
  }: HeaderManagerOptions) {
    this.sheet = sheet;
    this.rooms = rooms;
    this.fromRow = fromRow;
    this.fromColumn = fromColumn;
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

    this.sheet
      .getRange(this.fromRow, this.fromColumn, 1, this.rooms.length)
      .setValues([this.rooms.map(room => room.zh.name)]);
  }
}
