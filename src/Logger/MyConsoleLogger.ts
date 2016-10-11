import { LogListener, LogEntry, Logger } from "@agileis/sp-pnp-js/lib/utils/logging";

export class MyConsoleLogger implements LogListener {
    private _logPrefix: string;

    constructor(logPrefix?: string) {
        this._logPrefix = logPrefix ? logPrefix : "";
    }

    public log(entry: LogEntry) {
        let now = new Date();
        let log = `${this._logPrefix}# ${now.getDay()}.${now.getMonth()}.${now.getFullYear()}  ${now.getHours()}:${now.getMinutes()}:${now.getMilliseconds()}: ${entry.message}`;

        if (entry.data) {
            log += `- ${entry.data}`;
        }

        if (entry.level === Logger.LogLevel.Info || entry.level === Logger.LogLevel.Verbose) {
            console.log("\x1b[0m", log);
        } else if (entry.level === Logger.LogLevel.Warning) {
            console.warn("\x1b[33m", log);
        } else if (entry.level === Logger.LogLevel.Error) {
            console.error("\x1b[31m", log);
        }
    }
}
