/* eslint-disable @typescript-eslint/no-unused-vars */
import ENVIRONMENT from "../config";
import ImportantSessionsSheetManager from "./ImportantSessionsSheetManager";
import DataManager from "./DataManager";
import SessionSheetManager from "./SessionSheetManager";

global.entrypoint = function (): void {
  const dataManager = new DataManager(ENVIRONMENT.DATA_SOURCE);
  const importantSessionsSheet = new ImportantSessionsSheetManager(
    dataManager.data,
  );
  Logger.log("Dates: %s", dataManager.dates);
  Logger.log("Start: %s", dataManager.startHour);
  Logger.log("End: %s", dataManager.endHour);

  dataManager.dates.forEach((date, index) => {
    Logger.log("Init Day %s (%s)", index + 1, date);
    const sessionSheet = new SessionSheetManager(
      `Day ${index + 1} (${date})`,
      date,
      dataManager.data,
      importantSessionsSheet,
      dataManager.startHour,
      dataManager.endHour,
    );
    sessionSheet.fillData();
  });
};
