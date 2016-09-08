
import {Logger, LogListener, LogEntry} from "sp-pnp-js/lib/utils/logging";

export class MyConsoleLogger implements LogListener {
    log(entry: LogEntry) {
        console.log(entry.data + " - " + entry.level + " - " + entry.message);
    }
}