interface ColumnSchema {
  /** Title of the column */
  title: string;
  /** Column width */
  width: number;
  /**
   * Data setter when giving session detail
   * @param cell cell to be set
   * @param session session of the row
   */
  dataSetter: (
    cell: GoogleAppsScript.Spreadsheet.Range,
    session: EventSession,
  ) => void;
}
