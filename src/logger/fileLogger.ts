import { LogListener, LogEntry, Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import * as fs from "fs";

export class FileLogger implements LogListener {
    private static file: string;
    private logPrefix: string;

    constructor(logPrefix?: string) {
        this.logPrefix = logPrefix ? logPrefix + "# " : "";
        if (!FileLogger.file) {
            let now = new Date();
            FileLogger.file = `AISDEPLOY-${now.getFullYear()}${now.getMonth()}${now.getDay()}-${now.getHours()}${now.getMinutes()}.log`;
        }
    }

    public log(entry: LogEntry) {
        let now = new Date();
        let log = `[${Logger.LogLevel[entry.level]}] ${this.logPrefix}` +
            `${now.getDay()}.${now.getMonth()}.${now.getFullYear()}  ${now.getHours()}:${now.getMinutes()}:${now.getMilliseconds()}: ${entry.message}`;
        if (entry.data) {
            log += `- ${entry.data}`;
        }

        fs.appendFile(FileLogger.file, log + "\n", "utf8", error => {
            if (error) {
                console.error("\x1b[31m", error);
            }
        });
    }
}
