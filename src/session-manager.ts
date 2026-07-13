import { toMapById } from "./common";
import { CreateSessionDeps, Session } from "./session";

export class SessionManager {
  /**
   * @example ["2026-08-08", "2026-08-09"]
   */
  public readonly dates: string[];
  public readonly sessionsByDate: Record<string, DailySessionManager>;
  public readonly speakersById: Record<EventSpeakerId, EventSpeaker>;
  public readonly roomsById: Record<EventRoomId, EventRoom>;
  public readonly typesById: Record<EventSessionTypeId, EventSessionType>;

  public constructor(raw: EventData) {
    const { sessionsByDate, activeRooms, activeSpeakers, activeTypes } =
      formatSessions(
        {
          speakersById: toMapById(raw.speakers),
          roomsById: toMapById(raw.rooms),
          typesById: toMapById(raw.session_types),
        },
        raw.sessions,
      );
    this.sessionsByDate = sessionsByDate;
    this.roomsById = Object.fromEntries(activeRooms);
    this.speakersById = Object.fromEntries(activeSpeakers);
    this.typesById = Object.fromEntries(activeTypes);

    this.dates = Object.keys(this.sessionsByDate).sort((a, b) => {
      const dateA = new Date(a);
      const dateB = new Date(b);
      return dateA.getTime() - dateB.getTime();
    });
  }
}

export class DailySessionManager {
  public sessions: Session[] = [];
  public startsAt!: Date;
  public endsAt!: Date;

  public add(session: Session): void {
    this.sessions.push(session);

    if (!this.startsAt || session.startsAt < this.startsAt)
      this.startsAt = session.startsAt;
    if (!this.endsAt || session.endsAt > this.endsAt)
      this.endsAt = session.endsAt;
  }
}

function formatSessions(deps: CreateSessionDeps, rawSessions: EventSession[]) {
  const sessionsByDate: Record<string, DailySessionManager> = {};
  /** rooms that are really used by sessions */
  const activeRooms = new Map<EventRoomId, EventRoom>();
  /** speakers that are really hosting sessions */
  const activeSpeakers = new Map<EventSpeakerId, EventSpeaker>();
  /** types that are really used by sessions */
  const activeTypes = new Map<EventSessionTypeId, EventSessionType>();

  for (const rawSession of rawSessions) {
    if (!rawSession.start || !rawSession.end) {
      Logger.log(
        `WARN: Session ${rawSession.id} has no start or end time, skipping it.`,
      );
      continue;
    }

    const sessionInstance = new Session(deps, rawSession);

    if (sessionInstance.room)
      activeRooms.set(sessionInstance.room.id, sessionInstance.room);
    if (sessionInstance.type)
      activeTypes.set(sessionInstance.type.id, sessionInstance.type);
    for (const speaker of sessionInstance.speakers)
      activeSpeakers.set(speaker.id, speaker);

    const date = sessionInstance.date;
    const dailySessionManager = (sessionsByDate[date] ??=
      new DailySessionManager());
    dailySessionManager.add(sessionInstance);
  }

  return { sessionsByDate, activeRooms, activeSpeakers, activeTypes };
}
