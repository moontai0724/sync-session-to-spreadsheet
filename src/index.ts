/* eslint-disable @typescript-eslint/no-unused-vars */
import ENVIRONMENT from "../config";
import ImportantSessionsSheetManager from "./ImportantSessionsSheetManager";
import SessionManager from "./SessionManager";

global.entrypoint = function (): void {
  const sessionManager = new SessionManager(ENVIRONMENT.DATA_SOURCE);
  const importantSessionsSheet = new ImportantSessionsSheetManager();
};
