interface EventSpeaker {
  id: EventSpeakerId;
  avatar: string;
  zh: EventSpeakerDetail;
  en: EventSpeakerDetail;
}

interface EventSpeakerDetail {
  name: string;
  bio: string;
}

type EventSpeakerId = string;
