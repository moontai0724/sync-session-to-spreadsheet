export default class ImportantSessionsSheetManager {
  public readonly SHEET_NAME = "重要議程";
  public readonly PRESERVED_ROWS = 10;
  public readonly SCHEMA = [
    {
      title: "議程代號",
      width: 100,
      dataSetter: (
        cell: GoogleAppsScript.Spreadsheet.Range,
        session: EventSession,
      ): void => {
        const richValue = SpreadsheetApp.newRichTextValue()
          .setText(session.id)
          .setLinkUrl(session.uri)
          .build();

        cell.setRichTextValue(richValue);
      },
    },
    {
      title: "名稱",
      width: 500,
      dataSetter: (
        cell: GoogleAppsScript.Spreadsheet.Range,
        session: EventSession,
      ): void => {
        cell.setValue(session.zh.title);
      },
    },
    {
      title: "演講廳",
      width: 100,
      dataSetter: (
        cell: GoogleAppsScript.Spreadsheet.Range,
        session: EventSession,
      ): void => {
        cell.setValue(session.room);
      },
    },
    {
      title: "時間",
      width: 200,
      dataSetter: (
        cell: GoogleAppsScript.Spreadsheet.Range,
        session: EventSession,
      ): void => {
        const start = new Date(session.start);
        const end = new Date(session.end);
        const date = start.toLocaleDateString("zh-TW", {
          month: "2-digit",
          day: "2-digit",
        });
        const startTime = start.toLocaleTimeString("zh-TW", {
          hour: "2-digit",
          minute: "2-digit",
          hourCycle: "h24",
        });
        const endTime = end.toLocaleTimeString("zh-TW", {
          hour: "2-digit",
          minute: "2-digit",
          hourCycle: "h24",
        });
        const time = `${startTime} ~ ${endTime}`;
        cell.setValue(`${date} ${time}`);
      },
    },
  ];
  public spreadsheet;
  public sheet;

  public constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!this.isSheetExist()) {
      this.sheet = this.createSheet();
    } else {
      this.sheet = this.spreadsheet.getSheetByName(
        this.SHEET_NAME,
      ) as GoogleAppsScript.Spreadsheet.Sheet;
    }
  }

  public isSheetExist(): boolean {
    const sheet = this.spreadsheet.getSheetByName(this.SHEET_NAME);
    return !!sheet;
  }

  public createSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const sheet = this.spreadsheet.insertSheet(this.SHEET_NAME);

    const titles = this.SCHEMA.map(v => v.title);
    const titleRow = sheet.getRange(1, 1, 1, titles.length);
    titleRow.setValues([titles]);
    sheet.setFrozenRows(1);

    this.SCHEMA.forEach((value, index) => {
      sheet.setColumnWidth(index + 1, value.width);
    });

    sheet.deleteColumns(
      titles.length + 1,
      sheet.getMaxColumns() - titles.length,
    );
    sheet.deleteRows(
      this.PRESERVED_ROWS,
      sheet.getMaxRows() - this.PRESERVED_ROWS,
    );

    sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    return sheet;
  }

  public getIdColumn(): GoogleAppsScript.Spreadsheet.Range {
    const column = this.sheet.getRange(1, 1, this.sheet.getMaxRows(), 1);

    return column;
  }

  public getIds(): string[] {
    const ids = this.getIdColumn().getValues().flat().filter(Boolean);

    return ids;
  }

  public setDetails(session: EventSession): void {
    const rowIndex = this.getRowIndex(session.id);
    this.SCHEMA.forEach((schema, index) => {
      const cell = this.sheet.getRange(rowIndex, index + 1, 1, 1);
      schema.dataSetter(cell, session);
    });
  }

  public getRowIndex(id: string): number {
    const column = this.getIdColumn();
    const rowIndex = column.createTextFinder(id).findNext()?.getRow();
    if (!rowIndex) throw new Error(`ImportantSession: ${id} not found`);
    return rowIndex;
  }
}
