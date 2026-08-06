import { SessionManager } from "./data-manager";
import ENVIRONMENT from "../config";
import { SessionSheetManager } from "./session-sheet/sheet";

global.entrypoint = function (): void {
  const eventData = JSON.parse(
    UrlFetchApp.fetch(ENVIRONMENT.DATA_SOURCE).getContentText(),
  ) as EventData;
  const sessionManager = new SessionManager(eventData);
  const dates = Object.keys(sessionManager.sessionsByDate).sort();
  Logger.log("Dates: %s", dates);

  dates.forEach((date, day) => {
    const dailySessionManager = sessionManager.sessionsByDate[date];
    Logger.log(
      "Day %d: Date: %s, Sessions: %d, Starts at: %s, Ends at: %s",
      day + 1,
      date,
      dailySessionManager.sessions.length,
      dailySessionManager.startsAt.toISOString(),
      dailySessionManager.endsAt.toISOString(),
    );

    new SessionSheetManager(day + 1, dailySessionManager);
  });
};
