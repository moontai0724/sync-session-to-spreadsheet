interface EventItem {
  id: EventItemId;
  zh: EventItemDetail;
  en: EventItemDetail;
}

interface EventItemDetail {
  name: string;
}

type EventItemId = string;
