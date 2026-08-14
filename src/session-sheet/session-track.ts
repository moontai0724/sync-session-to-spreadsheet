import type { Session } from "../data-manager/session";
import type { EventDayManager } from "../data-manager/event-day";
import {
  SESSION_PRIORITY_COLORS,
  SPECIAL_SESSION_BORDER_COLOR,
  type SessionMarker,
} from "../marker-sheet";
import type { TimeManager } from "./time-slot";

export interface SessionTrackManagerOptions {
  room: EventRoom;
  column: number;
  timeManager: TimeManager;
  markers: ReadonlyMap<EventSessionId, SessionMarker>;
  eventDayManager: EventDayManager;
}

export class SessionTrackManager {
  public readonly room: EventRoom;
  public readonly column: number;
  public readonly timeManager: TimeManager;
  public readonly markers: ReadonlyMap<EventSessionId, SessionMarker>;
  public readonly eventDayManager: EventDayManager;
  private readonly sessions: Session[] = [];

  public constructor({
    room,
    column,
    timeManager,
    markers,
    eventDayManager,
  }: SessionTrackManagerOptions) {
    this.room = room;
    this.column = column;
    this.timeManager = timeManager;
    this.markers = markers;
    this.eventDayManager = eventDayManager;
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
      .map(({ session }) => ({
        session,
        startRow: this.timeManager.getRowByTime(session.startsAt),
        naturalEndRow: this.timeManager.getRowByTime(session.endsAt, true),
      }));

    sessions.forEach((current, index) => {
      const { session, startRow, naturalEndRow } = current;
      const next = sessions[index + 1];
      const hasSlotConflict =
        next !== undefined && next.startRow < naturalEndRow;
      const endRow = hasSlotConflict ? next.startRow : naturalEndRow;

      if (next && hasSlotConflict) {
        const hasRawOverlap = next.session.startsAt < session.endsAt;
        const rawOverlapEnd =
          next.session.endsAt < session.endsAt
            ? next.session.endsAt
            : session.endsAt;
        const rawOverlap = hasRawOverlap
          ? `${next.session.startsAt.toISOString()}–${rawOverlapEnd.toISOString()}`
          : "none (rounding-induced slot conflict)";
        Logger.log(
          `WARN: Session conflict in room ${this.room.id}: ` +
            `current=${session.id}, next=${next.session.id}, ` +
            `raw overlap=${rawOverlap}, ` +
            `slot overlap=rows ${next.startRow}–${naturalEndRow}, ` +
            `current truncated end=${next.session.startsAt.toISOString()} ` +
            `(row ${next.startRow}).`,
        );
      }

      if (endRow <= startRow) {
        const effectiveEndAt = hasSlotConflict
          ? next.session.startsAt
          : session.endsAt;
        Logger.log(
          `WARN: Session ${session.id} in room ${this.room.id} skipped: ` +
            `no visible slot after rounding/conflict truncation ` +
            `(${session.startsAt.toISOString()}–${effectiveEndAt.toISOString()}, ` +
            `rows ${startRow}–${endRow}).`,
        );
        return;
      }

      this.renderSession(session, startRow, endRow);
    });
  }

  private renderSession(
    session: Session,
    startRow: number,
    endRow: number,
  ): void {
    const richValue = SpreadsheetApp.newRichTextValue()
      .setText(formatSession(session))
      .setLinkUrl(this.eventDayManager.getSessionUrl(session.id))
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
        marker?.special ? SPECIAL_SESSION_BORDER_COLOR : "black",
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
