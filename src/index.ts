/* eslint-disable @typescript-eslint/no-unused-vars */
import ENVIRONMENT from "../config";
import ImportantSessionsSheetManager from "./ImportantSessionsSheetManager";
import DataManager from "./DataManager";

global.entrypoint = function (): void {
  const dtaManager = new DataManager(ENVIRONMENT.DATA_SOURCE);
  const importantSessionsSheet = new ImportantSessionsSheetManager();
};
