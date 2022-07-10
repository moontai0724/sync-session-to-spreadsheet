interface EventSession {
  /** id of the session */
  id: EventSessionId;
  /** type of the session */
  type: EventSessionTypeId;
  /** id of the room which session will be hold */
  room: EventRoomId;
  /** an ISO string represents start time of the session */
  start: string;
  /** an ISO string represents end time of the session  */
  end: string;
  /** language that will be spoke by speaker in this session */
  language: string;
  /** chinese title and description about this session */
  zh: EventSessionDetail;
  /** english title and description about this session */
  en: EventSessionDetail;
  /** speakers of this session */
  speakers: EventSpeakerId[];
  /** tags about this session */
  tags: EventTagId[];
  /** a link to hackmd co-writing note */
  co_write: string | null;
  /** a link to realtime qa */
  qa: string | null;
  /** a link to session slide */
  slide: string | null;
  /** a link to record video */
  record: string | null;
  /** a link to session on official website */
  uri: string;
}

interface EventSessionDetail {
  title: string;
  description: string;
}

type EventSessionId = string;
