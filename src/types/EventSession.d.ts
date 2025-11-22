interface EventSession {
  /** id of the session */
  id: EventSessionId;
  title_zh: string;
  description_zh: string;
  /** id of the room which session will be hold */
  room: EventRoomId;
  /** an ISO string represents start time of the session */
  start: string;
  /** an ISO string represents end time of the session  */
  end: string;
  /** a link to realtime qa */
  qa: string | null;
}

interface EventSessionDetail {
  title: string;
  description: string;
}

type EventSessionId = string;
