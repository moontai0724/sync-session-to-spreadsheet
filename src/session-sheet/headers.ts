const ROOM_NOTE_PREFIX = "sync-session-room:v1:";

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

  private readonly roomColumns = new Map<EventRoomId, number>();
  private managedRoomColumns: number[] = [];

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
    this.ensureHeaderRow();
    this.scanRoomMarkers();
    this.addMissingRoomColumns();

    const activeRoomIds = new Set(this.rooms.map(room => room.id));
    for (const [roomId, column] of Array.from(this.roomColumns.entries())) {
      if (activeRoomIds.has(roomId)) continue;
      this.sheet.getRange(this.fromRow, column).clearContent();
    }

    for (const room of this.rooms) {
      const column = this.getColumnByRoomId(room.id);
      this.sheet
        .getRange(this.fromRow, column)
        .setValue(room.zh.name)
        .setNote(`${ROOM_NOTE_PREFIX}${room.id}`);
      this.sheet.setColumnWidth(column, this.columnWidth);
    }

    this.managedRoomColumns = Array.from(this.roomColumns.values()).sort(
      (a, b) => a - b,
    );
  }

  public getColumnByRoomId(roomId: EventRoomId): number {
    const column = this.roomColumns.get(roomId);
    if (!column) throw new RangeError(`Room ${roomId} is not rendered`);
    return column;
  }

  public getManagedRoomColumns(): readonly number[] {
    return this.managedRoomColumns;
  }

  private ensureHeaderRow(): void {
    const missingRows = this.fromRow - this.sheet.getMaxRows();
    if (missingRows > 0) {
      this.sheet.insertRowsAfter(this.sheet.getMaxRows(), missingRows);
    }
  }

  private scanRoomMarkers(): void {
    this.roomColumns.clear();
    const notes = this.sheet
      .getRange(this.fromRow, 1, 1, this.sheet.getMaxColumns())
      .getNotes()[0];

    notes.forEach((note, offset) => {
      const column = offset + 1;
      if (column < this.fromColumn || !note.startsWith(ROOM_NOTE_PREFIX)) {
        return;
      }

      const roomId = note.slice(ROOM_NOTE_PREFIX.length);
      if (!roomId) return;
      if (this.roomColumns.has(roomId)) {
        throw new RangeError(`Duplicate room marker for ${roomId}`);
      }
      this.roomColumns.set(roomId, column);
    });
  }

  private addMissingRoomColumns(): void {
    const missingRooms = this.rooms.filter(
      room => !this.roomColumns.has(room.id),
    );
    if (missingRooms.length === 0) return;

    const insertAfter = Math.max(
      this.fromColumn - 1,
      this.sheet.getLastColumn(),
      ...Array.from(this.roomColumns.values()),
    );
    this.sheet.insertColumnsAfter(insertAfter, missingRooms.length);

    missingRooms.forEach((room, offset) => {
      this.roomColumns.set(room.id, insertAfter + offset + 1);
    });
  }
}
