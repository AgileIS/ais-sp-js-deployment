/// <reference path="../typings/index.d.ts" />

import {Logger, LogListener, LogEntry} from "sp-pnp-js/lib/utils/logging";

import {ISPObjectHandler} from "./interface/ObjectHandler/ispobjecthandler";
import {SiteHandler} from "./ObjectHandler/SiteHandler";
import {ListHandler} from "./ObjectHandler/ListHandler";
import {FieldHandler} from "./ObjectHandler/FieldHandler";
import {initAuth} from "./lib/initClient";
import * as fetch from "node-fetch";

class MyConsoleLogger implements LogListener {
    log(entry: LogEntry) {
        console.log(entry.data + " - " + entry.level + " - " + entry.message);
    }
}


Logger.subscribe(new MyConsoleLogger())
Logger.activeLogLevel = Logger.LogLevel.Verbose;

let fs = require('fs')
let args = require("minimist")(process.argv.slice(2));
Logger.write("start deployment script", Logger.LogLevel.Info);
Logger.write(JSON.stringify(args), 0);

if (args.f && args.p) {
    let config = JSON.parse(fs.readFileSync(args.f, 'utf8'));
    if (config.Url && config.User) {
        initAuth(config.Url, config.User, args.p);
        let siteUrl = config.Url;
        Logger.write(JSON.stringify(config), 0);
        let promise = chooseAndUseHandler(config, siteUrl);
        promise.then(() => {
            Logger.write("All Elements created");
        },
            (error) => {
                Logger.write("Error occured while creating Elemets - " + error);
            });
    }
}

export function resolveObjectHandler(key: string): ISPObjectHandler {
    switch (key) {
        case "Site":
            return new SiteHandler();
        case "List":
            return new ListHandler();
        case "Field":
            return new FieldHandler();
        default:
            break;
    }
}

export function chooseAndUseHandler(config: any, siteUrl: string) {
    return new Promise<boolean>((resolve, reject) => {
        let promiseArray = [];
        Object.keys(config).forEach((value, index) => {
            Logger.write("found config key " + value + " at index " + index, 0);
            let handler = resolveObjectHandler(value);
            if (typeof handler !== "undefined") {
                if (config[value] instanceof Array) {
                    let prom = Promise.resolve();
                    config[value].forEach(element => {
                        promiseArray.push(new Promise((resolve, reject) => {
                            prom = prom.then(() => {
                                let promise = handler.execute(element, siteUrl);
                                return promise;
                            }, (error) => {
                                return Promise.reject(error);
                            }).then((resolvedPromise) => {
                                Logger.write("Resolved Promise: " + JSON.stringify(resolvedPromise), 0);
                                chooseAndUseHandler(resolvedPromise, siteUrl).then(() => {
                                    resolve();
                                }, (error) => {
                                    reject(error);
                                });
                            }, (error) => {
                                reject(error);
                                Logger.write("Rejected: " + error, 0);
                                return null;
                            });
                        }));
                    });
                } else {
                    promiseArray.push(new Promise((resolve, reject) => {
                        let promise = handler.execute(config[value], siteUrl).then((resolvedPromise) => {
                            Logger.write("Resolved Promise: " + JSON.stringify(resolvedPromise), 0);
                            chooseAndUseHandler(resolvedPromise, siteUrl).then(
                                () => {
                                    resolve();
                                }, (error) => {
                                    reject(error);
                                });
                        }, (error) => {
                            reject(error);
                            Logger.write("Rejected: " + error, 0);
                        });
                    }));
                }
            }
        });
        Promise.all(promiseArray).then(() => {
            Logger.write("All Promises resolved", 0);
            resolve(true);
        }, (error) => {
            Logger.write("Not all Promises resolved - " + error, 0);
            reject(error);
        });
    });
}
