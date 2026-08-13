import { toMapById } from "../common";
import { DailySessionManager } from "./daily";
import type { EventDayMetadata } from "./event-day";
import { type CreateSessionDeps, Session } from "./session";

export class SessionManager {
  public readonly eventDayMetadata: EventDayMetadata;

  private readonly dailySessionsByDate = new Map<string, DailySessionManager>();
  private readonly sessionsById = new Map<EventSessionId, Session>();

  public constructor(
    { sessions, speakers, rooms, session_types: sessionTypes }: EventData,
    hiddenSessionIds: ReadonlySet<EventSessionId>,
    timezone: string,
  ) {
    const deps: CreateSessionDeps = {
      speakersById: toMapById(speakers),
      roomsById: toMapById(rooms),
      typesById: toMapById(sessionTypes),
    };
    const dates = new Set<string>();
    const dateBySessionId = new Map<EventSessionId, string>();
    const rawUrlBySessionId = new Map<EventSessionId, string>();
    const sessionIds = new Set<EventSessionId>();

    for (const rawSession of sessions) {
      if (sessionIds.has(rawSession.id)) {
        throw new RangeError(`Duplicate session ID: ${rawSession.id}`);
      }
      sessionIds.add(rawSession.id);

      const startsAt = new Date(rawSession.start);
      const endsAt = new Date(rawSession.end);
      const date = isValidDate(startsAt)
        ? getEventDate(startsAt, timezone)
        : undefined;
      if (date) {
        dates.add(date);
        dateBySessionId.set(rawSession.id, date);
      }

      const uri = rawSession.uri?.trim();
      if (uri) rawUrlBySessionId.set(rawSession.id, uri);

      const session = new Session(deps, rawSession);
      this.sessionsById.set(session.id, session);

      if (hiddenSessionIds.has(rawSession.id)) continue;
      if (!date || !isValidDate(endsAt) || endsAt <= startsAt) {
        Logger.log(
          `WARN: Session ${rawSession.id} has invalid time, skipping.`,
        );
        continue;
      }

      let dailySessionManager = this.dailySessionsByDate.get(date);
      if (!dailySessionManager) {
        dailySessionManager = new DailySessionManager();
        this.dailySessionsByDate.set(date, dailySessionManager);
      }
      dailySessionManager.add(session);
    }

    this.eventDayMetadata = {
      dates,
      dateBySessionId,
      rawUrlBySessionId,
      sessionIds,
    };
  }

  public getDailySessionManagerByDate(
    date: string,
  ): DailySessionManager | undefined {
    return this.dailySessionsByDate.get(date);
  }

  public getSession(sessionId: EventSessionId): Session | undefined {
    return this.sessionsById.get(sessionId);
  }
}

function getEventDate(date: Date, timezone: string): string {
  return Utilities.formatDate(date, timezone, "yyyy-MM-dd");
}

function isValidDate(date: Date): boolean {
  return !Number.isNaN(date.getTime());
}
