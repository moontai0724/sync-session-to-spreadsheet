import { drawSolidBorder } from "../draw";

const MINUTE_IN_MILLISECONDS = 60 * 1000;

export interface TimeManagerOptions {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  /** Header row of the time area. */
  fromRow: number;
  /** Inclusive start of the first time slot. */
  startAt: Date;
  /** Exclusive end, rounded up to the next whole hour. */
  endAt: Date;
  minutesPerUnit?: number;
  fromColumn?: number;
  headers?: readonly [string, string];
  columnWidth?: number;
  borderColor?: string;
}

/** Controls the two-column time area of a session sheet. */
export class TimeManager {
  public readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public readonly fromRow: number;
  public readonly fromColumn: number;
  public readonly startAt: Date;
  public readonly endAt: Date;
  public readonly minutesPerUnit: number;
  public readonly headers: readonly [string, string];
  public readonly columnWidth: number;
  public readonly borderColor: string;

  private readonly unitMilliseconds: number;
  private readonly slots: Date[];

  public constructor({
    sheet,
    fromRow,
    startAt,
    endAt,
    minutesPerUnit = 5,
    fromColumn = 1,
    headers = ["開始", "結束"],
    columnWidth = 50,
    borderColor = "black",
  }: TimeManagerOptions) {
    assertPositiveInteger("fromRow", fromRow);
    assertPositiveInteger("fromColumn", fromColumn);
    assertPositiveInteger("minutesPerUnit", minutesPerUnit);
    assertPositiveInteger("columnWidth", columnWidth);
    validateDate("startAt", startAt);
    validateDate("endAt", endAt);

    if (endAt <= startAt)
      throw new RangeError("endAt must be later than startAt");

    const unitMilliseconds = minutesPerUnit * MINUTE_IN_MILLISECONDS;
    const roundedStartAt = floorToUnit(startAt, minutesPerUnit);
    const roundedEndAt = ceilToHour(endAt);
    const duration = roundedEndAt.getTime() - roundedStartAt.getTime();
    if (duration % unitMilliseconds !== 0) {
      throw new RangeError(
        "The time range must be evenly divisible by minutesPerUnit",
      );
    }

    this.sheet = sheet;
    this.fromRow = fromRow;
    this.fromColumn = fromColumn;
    this.startAt = roundedStartAt;
    this.endAt = roundedEndAt;
    this.minutesPerUnit = minutesPerUnit;
    this.headers = headers;
    this.columnWidth = columnWidth;
    this.borderColor = borderColor;
    this.unitMilliseconds = unitMilliseconds;
    this.slots = this.createSlots();
  }

  public render(): void {
    const firstSlotRow = this.fromRow + 1;
    const lastSlotRow = this.fromRow + this.slots.length;

    ensureSheetSize(this.sheet, lastSlotRow, this.fromColumn + 1);
    clearStaleSlots(this.sheet, lastSlotRow, this.fromColumn);

    const headerRange = this.sheet.getRange(
      this.fromRow,
      this.fromColumn,
      1,
      2,
    );
    headerRange
      .setValues([[this.headers[0], this.headers[1]]])
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    this.sheet
      .getRange(firstSlotRow, this.fromColumn, this.slots.length, 2)
      .setValues(
        this.slots.map(startAt => [
          formatTime(startAt),
          formatTime(new Date(startAt.getTime() + this.unitMilliseconds)),
        ]),
      )
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    this.sheet.setColumnWidth(this.fromColumn, this.columnWidth);
    this.sheet.setColumnWidth(this.fromColumn + 1, this.columnWidth);

    for (const group of groupSlotsByHour(this.slots)) {
      drawSolidBorder(
        this.sheet.getRange(
          firstSlotRow + group.offset,
          this.fromColumn,
          group.length,
          2,
        ),
        this.borderColor,
      );
    }
  }

  public getRowByTime(time: Date, roundUp = false): number {
    validateDate("time", time);

    const offset = time.getTime() - this.startAt.getTime();
    if (offset < 0 || time > this.endAt) {
      throw new RangeError("time must be within the rendered time range");
    }
    const unitOffset = offset / this.unitMilliseconds;
    const slotOffset = roundUp ? Math.ceil(unitOffset) : Math.floor(unitOffset);
    return this.fromRow + 1 + slotOffset;
  }

  private createSlots(): Date[] {
    const slots: Date[] = [];
    for (
      let timestamp = this.startAt.getTime();
      timestamp < this.endAt.getTime();
      timestamp += this.unitMilliseconds
    ) {
      slots.push(new Date(timestamp));
    }
    return slots;
  }
}

function floorToUnit(date: Date, minutesPerUnit: number): Date {
  const result = new Date(date.getTime());
  result.setMinutes(
    Math.floor(result.getMinutes() / minutesPerUnit) * minutesPerUnit,
    0,
    0,
  );
  return result;
}

function ceilToHour(date: Date): Date {
  const result = new Date(date.getTime());
  if (result.getMinutes() > 0) {
    result.setHours(result.getHours() + 1, 0, 0, 0);
  }
  return result;
}

function validateDate(name: string, date: Date): void {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    throw new RangeError(`${name} must be a valid Date`);
  }
  if (date.getSeconds() !== 0 || date.getMilliseconds() !== 0) {
    throw new RangeError(`${name} must be aligned to a whole minute`);
  }
}

function ensureSheetSize(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  lastRow: number,
  lastColumn: number,
): void {
  const missingRows = lastRow - sheet.getMaxRows();
  if (missingRows > 0) sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);

  const missingColumns = lastColumn - sheet.getMaxColumns();
  if (missingColumns > 0)
    sheet.insertColumnsAfter(sheet.getMaxColumns(), missingColumns);
}

function clearStaleSlots(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  lastSlotRow: number,
  fromColumn: number,
): void {
  const staleRowCount = sheet.getLastRow() - lastSlotRow;
  if (staleRowCount <= 0) return;

  sheet
    .getRange(lastSlotRow + 1, fromColumn, staleRowCount, 2)
    .clearContent()
    .setBorder(false, false, false, false, false, false);
}

function groupSlotsByHour(
  slots: Date[],
): Array<{ offset: number; length: number }> {
  const groups: Array<{ offset: number; length: number }> = [];
  let previousHour = "";

  slots.forEach((slot, offset) => {
    const hour = `${slot.getFullYear()}-${slot.getMonth()}-${slot.getDate()}-${slot.getHours()}`;
    const current = groups[groups.length - 1];
    if (current && hour === previousHour) current.length++;
    else groups.push({ offset, length: 1 });
    previousHour = hour;
  });

  return groups;
}

function formatTime(date: Date): string {
  return `${date.getHours()}:${date.getMinutes().toString().padStart(2, "0")}`;
}

function assertPositiveInteger(name: string, value: number): void {
  if (!Number.isInteger(value) || value <= 0) {
    throw new RangeError(`${name} must be a positive integer`);
  }
}
