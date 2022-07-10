import ENVIRONMENT from "../config";
import SessionManager from "./SessionManager";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const sessionManager = new SessionManager(ENVIRONMENT.DATA_SOURCE);
