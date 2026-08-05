import { SessionManager } from "./data-manager";
import ENVIRONMENT from "../config";

global.entrypoint = function (): void {
  const eventData = JSON.parse(
    UrlFetchApp.fetch(ENVIRONMENT.DATA_SOURCE).getContentText(),
  ) as EventData;
  const sessionManager = new SessionManager(eventData);
  Logger.log("Dates: %s", sessionManager.dates);

  Object.entries(sessionManager.sessionsByDate).forEach(
    ([date, dailySessionManager], day) => {
      Logger.log(
        "Day %d: Date: %s, Sessions: %d, Starts at: %s, Ends at: %s",
        day + 1,
        date,
        dailySessionManager.sessions.length,
        dailySessionManager.startsAt.toISOString(),
        dailySessionManager.endsAt.toISOString(),
      );
    },
  );
};
