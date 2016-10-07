import { LogListener, LogEntry } from "@agileis/sp-pnp-js/lib/utils/logging";

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

        console.log(log);
    }
}
