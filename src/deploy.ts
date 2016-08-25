/// <reference path="../typings/index.d.ts" />

import {Logger, LogListener, LogEntry} from "sp-pnp-js/lib/utils/logging";

import {ISPObjectHandler} from "./interface/ObjectHandler/ispobjecthandler";
import {SiteHandler} from "./ObjectHandler/SiteHandler";
import {ListHandler} from "./ObjectHandler/ListHandler";
import {FieldHandler} from "./ObjectHandler/FieldHandler";
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
let promises = [];

if (args.f && args.p) {
    let config = JSON.parse(fs.readFileSync(args.f, 'utf8'));
    if (config.Url && config.User) {
        initAuth(config.Url, config.User, args.p);
        let siteUrl = config.Url;
        Logger.write(JSON.stringify(config), 0);
        chooseAndUseHandler(config, siteUrl);
        Promise.all(promises).then(() => {
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
    Object.keys(config).forEach((value, index) => {
        Logger.write("found config key " + value + " at index " + index, 0);
        let handler = resolveObjectHandler(value);
        if (typeof handler !== "undefined") {
            let prom = Promise.resolve();
            if (config[value] instanceof Array) {
                config[value].forEach(element => {
                    prom = prom.then((resolvedPromise) => {
                        let promi = handler.execute(element, siteUrl);
                        Logger.write("Resolved Promise: " + resolvedPromise);
                        if (typeof resolvedPromise !== "undefined") {
                          chooseAndUseHandler(resolvedPromise, siteUrl);
                        }
                        promises.push(promi);
                        return promi;
                    }).catch((error) => {
                        return error;
                    }) ;
                });
            } else {
                let promi = handler.execute(config[value], siteUrl).then((resolvedPromise) => {
                    Logger.write("Bla Resolved Promise: " + resolvedPromise);
                    if (typeof resolvedPromise !== "undefined") {
                      chooseAndUseHandler(resolvedPromise, siteUrl);
                    }
                });
                promises.push(promi);
            }
        }
    });
}



export interface Global {
    Headers: any;
    Request: any;
    fetch: any;
}

export function initAuth(url: string, username: string, password: string) {
    // Fixed missing Header & Request in node
    let fetch = require('node-fetch');
    global["Headers"] = fetch.Headers;
    global["Request"] = fetch.Request;
    global["fetch"] = fetch;

    // var httpNTLMClient = require("./lib/NTLMHttpClient");
    // httpNTLMClient.options.username = username;
    // httpNTLMClient.options.password = password;
    // httpNTLMClient.options.domain = domain;
    // httpNTLMClient.options.workstation = "";

    // var httpClient = require("agileis-sp-pnp-js/lib/net/HttpClient");
    // httpClient.HttpClient = httpNTLMClient.client;

    let httpBasicClient = require("./lib/BasicHttpClient");

    let userAndDommain = username.split("\\");

    httpBasicClient.options.username = userAndDommain[0] + "\\" + userAndDommain[1];
    httpBasicClient.options.password = password;

    let httpClient = require("sp-pnp-js/lib/net/HttpClient");
    httpClient.HttpClient = httpBasicClient.client;
}