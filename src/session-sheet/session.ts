import type { Session } from "../data-manager/session";
import type { HeaderManager } from "./headers";
import { SessionTrackManager } from "./session-track";
import type { TimeManager } from "./time-slot";

export interface SessionManagerOptions {
  sessions: readonly Session[];
  timeManager: TimeManager;
  headerManager: HeaderManager;
}

export class SessionManager {
  public readonly sessions: readonly Session[];
  public readonly timeManager: TimeManager;
  public readonly headerManager: HeaderManager;
  public readonly tracksByRoom: Map<EventRoomId, SessionTrackManager>;

  public constructor({
    sessions,
    timeManager,
    headerManager,
  }: SessionManagerOptions) {
    this.sessions = sessions;
    this.timeManager = timeManager;
    this.headerManager = headerManager;
    this.tracksByRoom = new Map(
      headerManager.rooms.map(
        room =>
          [
            room.id,
            new SessionTrackManager({
              room,
              column: headerManager.getColumnByRoomId(room.id),
              timeManager,
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
    for (const track of Array.from(this.tracksByRoom.values())) track.render();
  }
}
