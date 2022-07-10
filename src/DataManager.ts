export default class DataManager {
  public data: EventData;

  public constructor(public sourceUrl: string) {
    this.data = this.fetch();
  }

  public fetch(): EventData {
    const responsedText = UrlFetchApp.fetch(this.sourceUrl).getContentText();

    return JSON.parse(responsedText) as EventData;
  }
}
