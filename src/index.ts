import { EventDayManager, SessionManager } from "./data-manager";
import ENVIRONMENT from "../config";
import { MarkerSheetManager } from "./marker-sheet";
import { SessionSheetManager } from "./session-sheet/sheet";

global.entrypoint = function (): void {
  const eventData = JSON.parse(
    UrlFetchApp.fetch(ENVIRONMENT.DATA_SOURCE).getContentText(),
  ) as EventData;
  const markerSheetManager = new MarkerSheetManager();
  markerSheetManager.initialize();
  const sessionMarkers = markerSheetManager.getMarkers();
  const hiddenSessionIds = new Set(
    Array.from(sessionMarkers.entries())
      .filter(([, marker]) => marker.hidden)
      .map(([sessionId]) => sessionId),
  );

  const timezone = markerSheetManager.spreadsheet.getSpreadsheetTimeZone();
  const sessionManager = new SessionManager(
    eventData,
    hiddenSessionIds,
    timezone,
  );
  const eventDayManager = new EventDayManager(
    sessionManager.eventDayMetadata,
    ENVIRONMENT.SESSION_URL_TEMPLATE,
  );
  markerSheetManager.syncSessionDetails(sessionManager, eventDayManager);

  eventDayManager.days.forEach(({ day, date }) => {
    const dailySessionManager =
      sessionManager.getDailySessionManagerByDate(date);
    if (!dailySessionManager) return;
    new SessionSheetManager(
      day,
      dailySessionManager,
      sessionMarkers,
      eventDayManager,
    );
  });
};
