"use strict";

import { NTLM } from "./ntlm";
import * as http from "http";
import * as https from "https";
import * as url from "url";
import * as vm from "vm";
import { DeploymentConfig } from "./interface/Config/DeploymentConfig";
import { AuthenticationType } from "./Constants/AuthenticationType";

declare var hash: any;
declare var global: NodeJS.Global;
declare namespace NodeJS {
    interface Global {
        window: any;
    }
}

interface NodeJsomHandler {
    initialize(config: DeploymentConfig): Promise<void>;
}

interface NtlmOptions {
    domain: string;
    password: string;
    username: string;
    workstation: string;
}

class NodeJsomHandlerImpl implements NodeJsomHandler {
    public static instance: NodeJsomHandlerImpl;
    private static _agents: { [id: string]: http.Agent } = {};

    private _httpSavedRequest = undefined;
    private _httpsSavedRequest = undefined;
    private _authType: AuthenticationType;
    private _authOptions: string | NtlmOptions;

    public static httpRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        return NodeJsomHandlerImpl.instance._httpSavedRequest(NodeJsomHandlerImpl.setRequiredOptions(options), callback);
    }

    public static httpsRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        return NodeJsomHandlerImpl.instance._httpsSavedRequest(NodeJsomHandlerImpl.setRequiredOptions(options), callback);
    }

    private static setRequiredOptions(options: http.RequestOptions): http.RequestOptions {
        let requestOptions = undefined;
        if (typeof options === "string") {
            requestOptions = url.parse(options as string);
            requestOptions.url = options;
        } else {
            requestOptions = options;
        }

        requestOptions.headers = options.headers || {};
        requestOptions.headers.connection = "keep-alive";
        // requestOptions.agent = new http.Agent({ keepAlive: true, maxSockets: 1});

        return requestOptions;
    }

    public constructor() {
        NodeJsomHandlerImpl.instance = this;
    }

    public initialize(config: DeploymentConfig): Promise<void> {
        this._httpSavedRequest = http.request;
        http.request = NodeJsomHandlerImpl.httpRequest;
        this._httpsSavedRequest = https.request;
        https.request = NodeJsomHandlerImpl.httpsRequest;

        let promises = new Array<Promise<void>>();
        this._authType = config.User.authtype;

        if (this._authType === AuthenticationType.Ntlm) {
            this._authOptions = {
                domain: config.User.username.split("\\")[0],
                password: config.User.password,
                username: config.User.username.split("\\")[1],
                workstation: config.User.workstation ? config.User.workstation : "",
            };
            config.Sites.forEach(site => {
                promises.push(this.setupSiteContext(site.Url));
            });
        } else {
            this._authOptions = `Basic ${new Buffer(`${config.User.username}:${config.User.password}`).toString("base64")}`;
        }

        return new Promise<void>((resolve, reject) => {
            Promise.all(promises)
                .then(() => {
                    this.loadJsom(config.Sites[0].Url)
                        .then(() => { resolve(); })
                        .catch(error => { reject(error); });
                })
                .catch(error => { reject(error); });
        });
    }

    private setupSiteContext(siteUrl: string): Promise<void> {
        const lib = siteUrl.indexOf("https") > -1 ? https : http;
        let reqUrl = siteUrl.split("/").concat("_api/web/title".split("/")).join("/");
        let parsedUrl = url.parse(reqUrl as string);
        let authValue = this._authOptions;
        if (this._authType === AuthenticationType.Ntlm) {
            authValue = NTLM.createType1Message(this._authOptions);
            NodeJsomHandlerImpl._agents[siteUrl] = new lib.Agent({ keepAlive: true, maxSockets: 1 });
        }
        let options = {
            hostname: parsedUrl.hostname,
            path: parsedUrl.path,
            url: reqUrl,
            method: "GET",
            headers: {
                connection: "keep-alive",
                "content-type": "application/json;odata=verbose;charset=utf-8",
                "Authorization": authValue,
            },
            agent: NodeJsomHandlerImpl._agents[siteUrl],
        };

        return new Promise<void>((resolve, reject) => {
            http.get(options, firstResponse => {
                firstResponse.socket.emit("free");
                if (firstResponse.statusCode === 401) {
                    let type2msg = NTLM.parseType2Message(firstResponse.headers["www-authenticate"], error => {
                        reject(error);
                    });
                    let type3msg = NTLM.createType3Message(type2msg, this._authOptions);
                    options.headers.Authorization = type3msg;
                    http.get(options, secondResponse => {
                        firstResponse.socket.emit("free");
                        if (secondResponse.statusCode !== 200) {
                            reject(secondResponse.statusCode + ": after handshake!");
                        }
                        resolve();
                    });
                } else {
                    resolve();
                }
            }).on("error", error => reject(error));
        });
    }

    private loadJsom(siteUrl: string): Promise<void> {
        global.window = global;
        global.window.XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
        let urlObject = url.parse(siteUrl, true, true);

        let relativeUrl = urlObject.pathname;

        if (urlObject.search) {
            relativeUrl += urlObject.search;
        }

        if (urlObject.hash) {
            relativeUrl += hash;
        }

        global.window._spPageContextInfo = {
            webAbsoluteUrl: siteUrl,
            webServerRelativeUrl: relativeUrl,
        };

        global.window.navigator = {
            userAgent: "Node",
        };

        global.window.formdigest = {
            tagName: "INPUT",
            type: "hidden",
            value: "",
        };

        global.window.location = urlObject;

        global.window.document = {
            URL: window.location.href,
            cookie: "",
            documentElement: {},
            getElementsByName: function (name) {
                if (name === "__REQUESTDIGEST") {
                    return [global.window.formdigest];
                }
            },
            getElementsByTagName: function (name) {
                return [];
            },
        };

        let scripts = [
            "_layouts/15/init.debug.js",
            "_layouts/15/MicrosoftAjax.js",
            "_layouts/15/sp.core.debug.js",
            "_layouts/15/sp.runtime.debug.js",
            "_layouts/15/sp.debug.js",
        ];

        return new Promise<void>((resolve, reject) => {
            scripts.reduce((result, currentValue, currentIndex, array) => {
                return result.then((loadedScript: string) => {
                    if (loadedScript) {
                        vm.runInThisContext(loadedScript);
                    }
                    return new Promise((res, rej) => {
                        const lib = siteUrl.indexOf("https") > -1 ? https : http;
                        let combUrl = siteUrl.split("/").concat(currentValue.split("/")).join("/");
                        const request = lib.get(combUrl, response => {
                            const body = [];
                            response.on("data", (chunk) => body.push(chunk));
                            response.on("end", () => res(body.join("")));
                        });
                        request.on("error", (err) => reject(err));
                    });
                });
            }, Promise.resolve()).then((loadedScript: string) => {
                if (loadedScript) {
                    vm.runInThisContext(loadedScript);
                }
                resolve();
            });
        });
    }
}

export let NodeJsomHandler: NodeJsomHandler = new NodeJsomHandlerImpl();
