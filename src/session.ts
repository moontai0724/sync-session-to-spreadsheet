export interface CreateSessionDeps {
  speakersById: Partial<Record<EventSpeakerId, EventSpeaker>>;
  roomsById: Partial<Record<EventRoomId, EventRoom>>;
  typesById: Partial<Record<EventSessionTypeId, EventSessionType>>;
}

export class Session {
  public readonly id: string;
  public readonly title: string;
  public readonly date: string;
  public readonly startsAt: Date;
  public readonly endsAt: Date;
  public readonly url: string;

  public readonly room: EventRoom | null;
  public readonly type: EventSessionType | null;
  public readonly speakers: EventSpeaker[];

  public constructor(
    { speakersById, roomsById, typesById }: CreateSessionDeps,
    raw: EventSession,
  ) {
    this.id = raw.id;
    this.title = raw.zh.title;
    this.startsAt = new Date(raw.start);
    this.endsAt = new Date(raw.end);
    this.date = this.startsAt.toLocaleDateString("zh-TW", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });
    this.url = raw.uri;

    this.room = getRoom(roomsById, raw.room);
    this.type = getType(typesById, raw.type);
    this.speakers = getSpeakers(speakersById, raw.speakers);
  }
}

function getRoom(
  roomsById: Partial<Record<EventRoomId, EventRoom>>,
  roomId: EventRoomId,
): EventRoom | null {
  const room = roomsById[roomId];
  if (!room) {
    Logger.log(`WARN: Session has unknown room ${roomId}, skipping.`);
    return null;
  }
  return room;
}

function getType(
  typesById: Partial<Record<EventSessionTypeId, EventSessionType>>,
  typeId: EventSessionTypeId,
): EventSessionType | null {
  const type = typesById[typeId];
  if (!type) {
    Logger.log(`WARN: Session has unknown type ${typeId}, skipping.`);
    return null;
  }
  return type;
}

function getSpeakers(
  speakersById: Partial<Record<EventSpeakerId, EventSpeaker>>,
  speakerIds: EventSpeakerId[],
): EventSpeaker[] {
  return speakerIds
    .map(speakerId => {
      const speaker = speakersById[speakerId];
      if (!speaker) {
        Logger.log(`WARN: Session has unknown speaker ${speakerId}, skipping.`);
        return null;
      }
      return speaker;
    })
    .filter((v): v is EventSpeaker => !!v);
}
