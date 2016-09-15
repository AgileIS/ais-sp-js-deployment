/// <reference path="../typings/index.d.ts" />

import {Logger, LogListener, LogEntry} from "sp-pnp-js/lib/utils/logging";
import {ISPObjectHandler} from "./interface/ObjectHandler/ispobjecthandler";
import {SiteHandler} from "./ObjectHandler/SiteHandler";
import {ListHandler} from "./ObjectHandler/ListHandler";
import {FieldHandler} from "./ObjectHandler/FieldHandler";
import {ViewHandler} from "./ObjectHandler/ViewHandler";
import {ViewFieldHandler} from "./ObjectHandler/ViewFieldHandler";
import { ContentTypeHandler } from "./ObjectHandler/ContentTypeHandler";
import {HttpClient} from "./HttpClient/initClient";
import {MyConsoleLogger} from "./Logger/MyConsoleLogger";

let http = require("http");
let url = require("url");
let proxy = {
    protocol: "http:",
    hostname: "127.0.0.1",
    port: 8888,
};

let saveRequestObj = http.request;
http.request = function(options){
      if (typeof options === "string") { // options can be URL string.
            options = url.parse(options);
        }
        if (!options.host && !options.hostname) {
            throw new Error("host or hostname must have value.");
        }
        options.path = url.format(options.url);
        options.headers = options.headers || {};
        options.headers.Host = options.host || url.format({
            hostname: options.hostname,
            port: options.port
        });
        options.protocol = proxy.protocol;
        options.hostname = proxy.hostname;
        options.port = proxy.port;
        options.href = null;
        options.host = null;
        return saveRequestObj(options);
};


Logger.subscribe(new MyConsoleLogger());
Logger.activeLogLevel = Logger.LogLevel.Verbose;

let fs = require("fs");
let args = require("minimist")(process.argv.slice(2));
Logger.write("start deployment script", Logger.LogLevel.Info);
Logger.write(JSON.stringify(args), 0);

if (args.f && args.p) {
    let config = JSON.parse(fs.readFileSync(args.f, "utf8"));
    if (config.User) {
        HttpClient.initAuth(config.User, args.p);
        Logger.write(JSON.stringify(config), 0);
        Promise.all(chooseAndUseHandler(config, null)).then(() => {
            Logger.write("All Elements created", 1);
            return;
        }).catch((error) => {
            Logger.write("Error occured while creating Elemets - " + error, 1);
            return;
        });
    }
}
// todo: Factory auslagern mit execute und parentPromise.then => execute Handler ???
function resolveObjectHandler(key: string): ISPObjectHandler {
    switch (key) {
        case "Site":
            return new SiteHandler();
        case "List":
            return new ListHandler();
        case "Field":
            return new FieldHandler();
        case "View":
            return new ViewHandler();
        /* do we need this handler any more?*/
        case "ViewFields":
            return new ViewFieldHandler();
        case "ContentTypes":
            return new ContentTypeHandler();
        default:
            break;
    }
}

function promiseStatus(p) {
    return p.then(function (val) { return { status: "resolved", val: val }; },
        function (val) { return { status: "rejected", val: val }; }
    );
}

function chooseAndUseHandler(config: any, parent?: Promise<any>): Array<Promise<any>> {
    let promises: Array<Promise<any>> = [];

    Object.keys(config).forEach((value, index) => {
        Logger.write("found config key " + value + " at index " + index, 0);
        let handler = resolveObjectHandler(value);
        if (typeof handler !== "undefined") {
            Logger.write("found handler " + handler.constructor.name + " for config key " + value, 0);
            if (config[value] instanceof Array) {
                config[value].forEach(element => {
                    Logger.write("call object handler " + handler.constructor.name + " with element:" + JSON.stringify(element), 0);
                    let promise = handler.execute(element, parent);
                    promises.push(promise);
                    promises.concat(chooseAndUseHandler(element, promise));
                });
            }
            else {
                Logger.write("call object handler " + handler.constructor.name + " with element:" + JSON.stringify(config[value]), 0);
                let promise = handler.execute(config[value], parent);
                promises.push(promise);
                promises.concat(chooseAndUseHandler(config[value], promise));
            }
        }
    });
    return promises;
}
