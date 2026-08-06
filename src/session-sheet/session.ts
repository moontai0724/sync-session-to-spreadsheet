import type { Session } from "../data-manager/session";
import type { HeaderManager } from "./headers";
import type { TimeManager } from "./time-slot";

export interface SessionManagerOptions {
  sessions: readonly Session[];
  timeManager: TimeManager;
  headerManager: HeaderManager;
}

interface IndexedSession {
  session: Session;
  index: number;
}

export class SessionManager {
  public readonly sessions: readonly Session[];
  public readonly timeManager: TimeManager;
  public readonly headerManager: HeaderManager;

  public constructor({
    sessions,
    timeManager,
    headerManager,
  }: SessionManagerOptions) {
    this.sessions = sessions;
    this.timeManager = timeManager;
    this.headerManager = headerManager;
  }

  public render(): void {
    const sessionsByRoom = new Map<EventRoomId, IndexedSession[]>();

    this.sessions.forEach((session, index) => {
      if (!session.room) return;
      const roomSessions = sessionsByRoom.get(session.room.id) ?? [];
      roomSessions.push({ session, index });
      sessionsByRoom.set(session.room.id, roomSessions);
    });

    sessionsByRoom.forEach(roomSessions => this.renderRoom(roomSessions));
  }

  private renderRoom(roomSessions: IndexedSession[]): void {
    roomSessions.sort(
      (a, b) =>
        a.session.startsAt.getTime() - b.session.startsAt.getTime() ||
        a.index - b.index,
    );

    roomSessions.forEach(({ session }, index) => {
      const nextSession = roomSessions[index + 1]?.session;
      const effectiveEndAt =
        nextSession && nextSession.startsAt < session.endsAt
          ? nextSession.startsAt
          : session.endsAt;
      if (effectiveEndAt <= session.startsAt) return;

      this.renderSession(session, effectiveEndAt);
    });
  }

  private renderSession(session: Session, endAt: Date): void {
    if (!session.room) return;

    const startRow = this.timeManager.getRowByTime(session.startsAt);
    const endRow = this.timeManager.getRowByTime(endAt);
    const column = this.headerManager.getColumnByRoomId(session.room.id);
    const richValue = SpreadsheetApp.newRichTextValue()
      .setText(formatSession(session))
      .setLinkUrl(session.url)
      .build();

    this.timeManager.sheet
      .getRange(startRow, column, endRow - startRow, 1)
      .merge()
      .setRichTextValue(richValue)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true)
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        "black",
        SpreadsheetApp.BorderStyle.SOLID,
      );
  }
}

function formatSession(session: Session): string {
  const speakerNames = session.speakers
    .map(speaker => speaker.zh.name)
    .join("、");
  return `${session.title} by [${session.speakers.length}] ${speakerNames}`;
}
