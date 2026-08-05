import { Session } from "../session";

export class DailySessionManager {
  public sessions: Session[] = [];
  public startsAt!: Date;
  public endsAt!: Date;

  public activeRooms: Map<EventRoomId, EventRoom> = new Map();
  public activeSpeakers: Map<EventSpeakerId, EventSpeaker> = new Map();
  public activeTypes: Map<EventSessionTypeId, EventSessionType> = new Map();

  public add(session: Session): void {
    this.sessions.push(session);

    if (!this.startsAt || session.startsAt < this.startsAt)
      this.startsAt = session.startsAt;
    if (!this.endsAt || session.endsAt > this.endsAt)
      this.endsAt = session.endsAt;

    if (session.room) this.activeRooms.set(session.room.id, session.room);
    if (session.type) this.activeTypes.set(session.type.id, session.type);
    for (const speaker of session.speakers)
      this.activeSpeakers.set(speaker.id, speaker);
  }
}
