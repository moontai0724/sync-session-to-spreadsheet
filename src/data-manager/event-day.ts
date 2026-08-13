export interface EventDay {
  readonly day: number;
  readonly date: string;
}

export interface EventDayMetadata {
  readonly dates: ReadonlySet<string>;
  readonly dateBySessionId: ReadonlyMap<EventSessionId, string>;
  readonly rawUrlBySessionId: ReadonlyMap<EventSessionId, string>;
  readonly sessionIds: ReadonlySet<EventSessionId>;
}

export class EventDayManager {
  public readonly days: readonly EventDay[];
  private readonly dayByDate: ReadonlyMap<string, number>;

  public constructor(
    private readonly metadata: EventDayMetadata,
    private readonly sessionUrlTemplate?: string,
  ) {
    this.days = Array.from(metadata.dates)
      .sort()
      .map((date, index) => ({ day: index + 1, date }));
    this.dayByDate = new Map(this.days.map(({ date, day }) => [date, day]));
  }

  public getDayBySessionId(sessionId: EventSessionId): number {
    if (!this.metadata.sessionIds.has(sessionId)) {
      throw new RangeError(`Unknown session ${sessionId}`);
    }
    const date = this.metadata.dateBySessionId.get(sessionId);
    const day = date ? this.dayByDate.get(date) : undefined;
    if (!day) {
      throw new RangeError(
        `Cannot determine the event day for session ${sessionId}`,
      );
    }
    return day;
  }

  public getSessionUrl(sessionId: EventSessionId): string {
    const rawUrl = this.metadata.rawUrlBySessionId.get(sessionId);
    if (rawUrl) return rawUrl;
    if (!this.metadata.sessionIds.has(sessionId)) {
      throw new RangeError(`Unknown session ${sessionId}`);
    }

    const template = this.sessionUrlTemplate;
    if (!template) {
      throw new Error(
        `Session ${sessionId} has no URI and SESSION_URL_TEMPLATE is not configured`,
      );
    }

    const day = template.includes("{day}")
      ? this.getDayBySessionId(sessionId)
      : undefined;
    return applySessionUrlTemplate(template, sessionId, day);
  }
}

function applySessionUrlTemplate(
  template: string,
  sessionId: EventSessionId,
  day?: number,
): string {
  const withDay =
    day === undefined ? template : template.replace(/\{day\}/g, String(day));
  return withDay.replace(/\{id\}/g, sessionId);
}
