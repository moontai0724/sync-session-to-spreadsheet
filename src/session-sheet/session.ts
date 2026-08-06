import type { Session } from "../data-manager/session";
import type { SessionPriority } from "../marker-sheet";
import type { HeaderManager } from "./headers";
import { SessionTrackManager } from "./session-track";
import type { TimeManager } from "./time-slot";

export interface SessionManagerOptions {
  sessions: readonly Session[];
  timeManager: TimeManager;
  headerManager: HeaderManager;
  priorities: ReadonlyMap<EventSessionId, SessionPriority>;
}

export class SessionManager {
  public readonly sessions: readonly Session[];
  public readonly timeManager: TimeManager;
  public readonly headerManager: HeaderManager;
  public readonly priorities: ReadonlyMap<EventSessionId, SessionPriority>;
  public readonly tracksByRoom: Map<EventRoomId, SessionTrackManager>;

  public constructor({
    sessions,
    timeManager,
    headerManager,
    priorities,
  }: SessionManagerOptions) {
    this.sessions = sessions;
    this.timeManager = timeManager;
    this.headerManager = headerManager;
    this.priorities = priorities;
    this.tracksByRoom = new Map(
      headerManager.rooms.map(
        room =>
          [
            room.id,
            new SessionTrackManager({
              room,
              column: headerManager.getColumnByRoomId(room.id),
              timeManager,
              priorities,
            }),
          ] as const,
      ),
    );

    for (const session of sessions) {
      if (!session.room) {
        Logger.log(`Session ${session.id} has no room assigned, skipping`);
        continue;
      }

      const track = this.tracksByRoom.get(session.room.id);
      if (!track)
        throw new RangeError(`No track found for room ${session.room.id}`);

      track.add(session);
    }
  }

  public render(): void {
    this.clearManagedColumns();
    for (const track of Array.from(this.tracksByRoom.values())) track.render();
  }

  private clearManagedColumns(): void {
    const fromRow = this.headerManager.fromRow + 1;
    const rowCount =
      this.timeManager.sheet.getMaxRows() - this.headerManager.fromRow;
    if (rowCount <= 0) return;

    for (const group of groupColumns(
      this.headerManager.getManagedRoomColumns(),
    )) {
      this.timeManager.sheet
        .getRange(fromRow, group.fromColumn, rowCount, group.columnCount)
        .breakApart()
        .clear();
    }
  }
}

function groupColumns(
  columns: readonly number[],
): Array<{ fromColumn: number; columnCount: number }> {
  const groups: Array<{ fromColumn: number; columnCount: number }> = [];
  for (const column of columns) {
    const current = groups[groups.length - 1];
    if (current && current.fromColumn + current.columnCount === column) {
      current.columnCount++;
    } else {
      groups.push({ fromColumn: column, columnCount: 1 });
    }
  }
  return groups;
}
