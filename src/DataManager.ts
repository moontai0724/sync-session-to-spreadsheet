export default class DataManager {
  public data: EventData;
  public startHour: number;
  public endHour: number;
  public dates: string[];

  public constructor(public sourceUrl: string) {
    this.data = this.fetch();
    this.data.rooms = this.getActiveRooms();
    const [startHour, endHour] = this.getHourRange();
    this.startHour = startHour;
    this.endHour = endHour;
    this.dates = this.getDates();
  }

  public fetch(): EventData {
    const responsedText = UrlFetchApp.fetch(this.sourceUrl).getContentText();

    return JSON.parse(responsedText) as EventData;
  }

  public getActiveRooms(): EventRoom[] {
    const realRoomIds = Array.from(
      new Set(this.data.sessions.map(session => session.room).sort()),
    );
    const rooms = realRoomIds
      .map(roomId => this.data.rooms.find(room => room.id === roomId))
      .filter(Boolean) as EventRoom[];

    return rooms;
  }

  public getHourRange(): [number, number] {
    let min = 24;
    let max = 0;

    this.data.sessions.forEach(session => {
      const start = new Date(session.start).getHours();
      if (start < min) {
        min = start;
      }

      const end = new Date(session.end).getHours();
      if (end > max) {
        max = end;
      }
    });

    return [min, max];
  }

  public getDates(): string[] {
    const dates = Array.from(
      new Set(
        this.data.sessions.map(session => {
          const date = new Date(session.start);
          const dateString = date.toLocaleDateString("zh-TW", {
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
          });

          return dateString;
        }),
      ),
    ).sort((date1, date2) => {
      const date1Date = new Date(date1);
      const date2Date = new Date(date2);

      return date1Date.getTime() - date2Date.getTime();
    });

    return dates;
  }
}
