import { toMapById } from "../common";
import { type CreateSessionDeps, Session } from "./session";
import { DailySessionManager } from "./daily";

export class SessionManager {
  /**
   * @example ["2026-08-08", "2026-08-09"]
   */
  public readonly dates: string[];
  public readonly sessionsByDate: Record<string, DailySessionManager> = {};

  public readonly activeRooms: Map<EventRoomId, EventRoom> = new Map();
  public readonly activeSpeakers: Map<EventSpeakerId, EventSpeaker> = new Map();
  public readonly activeTypes: Map<EventSessionTypeId, EventSessionType> =
    new Map();

  public constructor({
    sessions,
    speakers,
    rooms,
    session_types: sessionTypes,
  }: EventData) {
    const deps: CreateSessionDeps = {
      speakersById: toMapById(speakers),
      roomsById: toMapById(rooms),
      typesById: toMapById(sessionTypes),
    };

    for (const rawSession of sessions) {
      if (!rawSession.start || !rawSession.end) {
        Logger.log(
          `WARN: Session ${rawSession.id} has no start or end time, skipping it.`,
        );
        continue;
      }

      const session = new Session(deps, rawSession);
      this.add(session);
    }

    this.dates = Object.keys(this.sessionsByDate).sort((a, b) => {
      const dateA = new Date(a);
      const dateB = new Date(b);
      return dateA.getTime() - dateB.getTime();
    });
  }

  public add(session: Session): void {
    const date = session.date;
    const dailySessionManager = (this.sessionsByDate[date] ??=
      new DailySessionManager());
    dailySessionManager.add(session);

    if (session.room) this.activeRooms.set(session.room.id, session.room);
    if (session.type) this.activeTypes.set(session.type.id, session.type);
    for (const speaker of session.speakers)
      this.activeSpeakers.set(speaker.id, speaker);
  }
}
