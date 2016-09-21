
import { LogListener, LogEntry } from "@agileis/sp-pnp-js/lib/utils/logging";

export class MyConsoleLogger implements LogListener {
    public log(entry: LogEntry) {
        let now = new Date();
        let log = `${now.getDay()}.${now.getMonth()}.${now.getFullYear()}  ${now.getHours()}:${now.getMinutes()}:${now.getMilliseconds()}: ${entry.message}`;

        if (entry.data) {
            log += `- ${entry.data}`;
        }

        console.log(log);
    }
}
