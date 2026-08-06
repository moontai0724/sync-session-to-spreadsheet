import type { Session } from "../data-manager/session";
import {
  SESSION_PRIORITY_COLORS,
  SPECIAL_SESSION_COLOR,
  type SessionMarker,
} from "../marker-sheet";
import type { TimeManager } from "./time-slot";

export interface SessionTrackManagerOptions {
  room: EventRoom;
  column: number;
  timeManager: TimeManager;
  markers: ReadonlyMap<EventSessionId, SessionMarker>;
}

export class SessionTrackManager {
  public readonly room: EventRoom;
  public readonly column: number;
  public readonly timeManager: TimeManager;
  public readonly markers: ReadonlyMap<EventSessionId, SessionMarker>;
  private readonly sessions: Session[] = [];

  public constructor({
    room,
    column,
    timeManager,
    markers,
  }: SessionTrackManagerOptions) {
    this.room = room;
    this.column = column;
    this.timeManager = timeManager;
    this.markers = markers;
  }

  public add(session: Session): void {
    if (session.room?.id !== this.room.id) {
      throw new RangeError(
        `Session ${session.id} does not belong to this track`,
      );
    }
    this.sessions.push(session);
  }

  public render(): void {
    const sessions = this.sessions
      .map((session, index) => ({ session, index }))
      .sort(
        (a, b) =>
          a.session.startsAt.getTime() - b.session.startsAt.getTime() ||
          a.index - b.index,
      )
      .map(({ session }) => session);

    sessions.forEach((session, index) => {
      const nextSession = sessions[index + 1];
      const effectiveEndAt =
        nextSession && nextSession.startsAt < session.endsAt
          ? nextSession.startsAt
          : session.endsAt;
      if (effectiveEndAt <= session.startsAt) return;

      this.renderSession(session, effectiveEndAt);
    });
  }

  private renderSession(session: Session, endAt: Date): void {
    const startRow = this.timeManager.getRowByTime(session.startsAt);
    const endRow = this.timeManager.getRowByTime(endAt);
    const richValue = SpreadsheetApp.newRichTextValue()
      .setText(formatSession(session))
      .setLinkUrl(session.url)
      .build();
    const marker = this.markers.get(session.id);

    const range = this.timeManager.sheet
      .getRange(startRow, this.column, endRow - startRow, 1)
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
        marker?.special ? SPECIAL_SESSION_COLOR : "black",
        marker?.special
          ? SpreadsheetApp.BorderStyle.SOLID_THICK
          : SpreadsheetApp.BorderStyle.SOLID,
      );

    if (marker?.priority) {
      range.setBackground(SESSION_PRIORITY_COLORS[marker.priority]);
    }
  }
}

function formatSession(session: Session): string {
  const speakerNames = session.speakers
    .map(speaker => speaker.zh.name)
    .join("、");
  return `${session.title} by [${session.speakers.length}] ${speakerNames}`;
}
