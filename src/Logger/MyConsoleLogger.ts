
import { LogListener, LogEntry } from "sp-pnp-js/lib/utils/logging";

export class MyConsoleLogger implements LogListener {
    public log(entry: LogEntry) {
        let log = entry.level + " - " + entry.message;

        if (entry.data) {
            log += " - " + entry.data;
        }

        console.log(log);
    }
}
