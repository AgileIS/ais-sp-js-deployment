/// <reference path="../typings/index.d.ts" />

import {Logger, LogListener, LogEntry} from "sp-pnp-js/lib/utils/logging";

import {ISPObjectHandler} from "./interface/ispobjecthandler";
import {SiteHandler} from "./ObjectHandler/SiteHandler"
import {ListHandler} from "./ObjectHandler/ListHandler"

class MyConsoleLogger implements LogListener{
    log(entry: LogEntry ){
        console.log(entry.data + " - " + entry.level + " - " +  entry.message);
    }
    
}

Logger.subscribe(new MyConsoleLogger())
Logger.activeLogLevel = Logger.LogLevel.Verbose;

let fs = require('fs')
let args = require("minimist")(process.argv.slice(2));
Logger.write("start deployment script", Logger.LogLevel.Info);
Logger.write(JSON.stringify(args), 0);

if(args.f){
    let config = JSON.parse(fs.readFileSync(args.f,'utf8'));
    Logger.write(JSON.stringify(config), 0);

    Object.keys(config).forEach((value, index) =>{
        Logger.write("found config key "+ value + " at index " + index,0);
        let handler = resolveObjectHandler(value);
        if(config[value] instanceof Array){
            config[value].forEach(element => {
                handler.execute(element);
            });    
        }else{
            handler.execute(config[value]);
        }
        
    });

}

export function resolveObjectHandler(key: string): ISPObjectHandler {
    switch (key) {
        case "Site":
            return new SiteHandler();
            case "List":
            return new ListHandler()
        default:
            break;
    }
}

