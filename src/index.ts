/* eslint-disable @typescript-eslint/no-unused-vars */
import ENVIRONMENT from "../config";
import ImportantSessionsSheetManager from "./ImportantSessionsSheetManager";
import DataManager from "./DataManager";
import SessionSheetManager from "./SessionSheetManager";

global.entrypoint = function (): void {
  const dataManager = new DataManager(ENVIRONMENT.DATA_SOURCE);
  const importantSessionsSheet = new ImportantSessionsSheetManager();
  Logger.log("Dates: %s", dataManager.dates);
  Logger.log("Start: %s", dataManager.startHour);
  Logger.log("End: %s", dataManager.endHour);

  dataManager.dates.forEach((date, index) => {
    Logger.log("Init Day %s (%s)", index, date);
    const sessionSheet = new SessionSheetManager(
      `Day ${index} (${date})`,
      date,
      dataManager.startHour,
      dataManager.endHour,
      dataManager.data,
    );
  });
};
