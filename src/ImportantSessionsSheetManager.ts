export default class ImportantSessionsSheetManager {
  /** Sheet name of important sessions, will auto create when construct if not exists. */
  public readonly SHEET_NAME = "重要議程";
  /** Default preserved empty rows for fill-in, no any side effect if changed. */
  public readonly PRESERVED_ROWS = 10;
  /** Schema of the sheet content by column */
  public readonly SCHEMA: ColumnSchema[] = [
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
  public sessions;

  /**
   * @param data Data to be used for fill-in session infos
   */
  public constructor(
    data: EventData = {
      sessions: [],
      speakers: [],
      session_types: [],
      rooms: [],
      tags: [],
    },
  ) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = this.spreadsheet.getSheetByName(this.SHEET_NAME);
    this.sheet = sheet ?? this.createSheet();
    const sessionIds = this.getIdColumn().getValues().flat().filter(Boolean);
    this.sessions = data.sessions.filter(session =>
      sessionIds.includes(session.id),
    );
    this.sessions.forEach(session => this.setDetails(session));
  }

  /**
   * Create a default structure of sheet with preserved rows by schema.
   * @returns Sheet of created important sessions
   */
  private createSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    Logger.log(`Create sheet: ${this.SHEET_NAME}`);
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

  /**
   * Get a column range of all important sessions' id.
   * @returns Column of ids exclude title row.
   */
  public getIdColumn(): GoogleAppsScript.Spreadsheet.Range {
    const column = this.sheet.getRange(2, 1, this.sheet.getMaxRows(), 1);

    return column;
  }

  /**
   * Set details of session in sheet by id.
   * @param session session to be set
   * @throws Error if session not found
   */
  public setDetails(session: EventSession): void {
    Logger.log(`Set details of important session: ${session.id}`);
    const rowIndex = this.getRowIndex(session.id);
    this.SCHEMA.forEach((schema, index) => {
      const cell = this.sheet.getRange(rowIndex, index + 1, 1, 1);
      Logger.log(
        `Set details of important session @ %s: %s %s`,
        cell.getA1Notation(),
        session.id,
        schema.title,
      );
      schema.dataSetter(cell, session);
    });
  }

  /**
   * Get row index of session by id.
   * @param sessionId id of the session to be found
   * @throws Error if session not found
   * @returns ㄙrow index of session
   */
  public getRowIndex(sessionId: string): number {
    const column = this.getIdColumn();
    const rowIndex = column.createTextFinder(sessionId).findNext()?.getRow();
    if (!rowIndex) throw new Error(`ImportantSession: ${sessionId} not found`);
    return rowIndex;
  }
}
