"use strict";

import { NTLM } from "./ntlm";
// import * as http from "http";
// import * as https from "https";
import http = require("http");
import https = require("https");
import * as url from "url";
import * as vm from "vm";
import { ISiteDeploymentConfig } from "./Interfaces/Config/SiteDeploymentConfig";
import { AuthenticationType } from "./Constants/AuthenticationType";
import { Util } from "./Util/Util";

declare var global: NodeJS.Global;
declare namespace NodeJS {
    interface Global { // tslint:disable-line
        window: any;
    }
}

interface INodeJsomHandler {
    initialize(siteDeploymentConfig: ISiteDeploymentConfig): Promise<void>;
}

interface INtlmOptions {
    domain: string;
    password: string;
    username: string;
    workstation: string;
}

class NodeJsomHandlerImpl implements INodeJsomHandler {
    public static instance: NodeJsomHandlerImpl;
    private static agents: { [id: string]: http.Agent } = {};
    private static authType: AuthenticationType;
    private static authOptions: string | INtlmOptions;

    private httpSavedRequest = undefined;
    private httpsSavedRequest = undefined;

    public static httpRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        if (typeof options !== "string" && !options.protocol) {
            options.protocol = "http:";
        }
        return NodeJsomHandlerImpl.instance.httpSavedRequest(NodeJsomHandlerImpl.setRequiredOptions(options), callback);
    }

    public static httpsRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        if (typeof options !== "string" && !options.protocol) {
            options.protocol = "https:";
        }
        return NodeJsomHandlerImpl.instance.httpsSavedRequest(NodeJsomHandlerImpl.setRequiredOptions(options), callback);
    }

    private static setRequiredOptions(options: http.RequestOptions): http.RequestOptions {
        let requestOptions = undefined;
        if (typeof options !== "string") {
            requestOptions = options;
            if (options.headers["User-Agent"] === "node-XMLHttpRequest" || options.headers["X-Request-With"] === "XMLHttpRequest") {
                requestOptions.headers = options.headers || {};
                requestOptions.headers.connection = "keep-alive";
                if (!requestOptions.url) {
                    requestOptions.url = Util.JoinAndNormalizeUrl([options.protocol, options.host, options.path]);
                }
                if (NodeJsomHandlerImpl.authType === AuthenticationType.NTLM) {
                    requestOptions.agent = NodeJsomHandlerImpl.agents[requestOptions.url.split("/_")[0]];
                } else {
                    requestOptions.headers.Authorization = NodeJsomHandlerImpl.authOptions;
                }
            }
        } else {
            requestOptions = options;
        }

        return requestOptions;
    }

    public constructor() {
        NodeJsomHandlerImpl.instance = this;
    }

    public initialize(siteDeploymentConfig: ISiteDeploymentConfig): Promise<void> {
        this.httpSavedRequest = http.request;
        http.request = NodeJsomHandlerImpl.httpRequest;
        this.httpsSavedRequest = https.request;
        https.request = NodeJsomHandlerImpl.httpsRequest;

        let promises = new Array<Promise<void>>();
        NodeJsomHandlerImpl.authType = siteDeploymentConfig.User.authtype;

        if (NodeJsomHandlerImpl.authType === AuthenticationType.NTLM) {
            NodeJsomHandlerImpl.authOptions = {
                domain: siteDeploymentConfig.User.username.split("\\")[0],
                password: siteDeploymentConfig.User.password,
                username: siteDeploymentConfig.User.username.split("\\")[1],
                workstation: siteDeploymentConfig.User.workstation ? siteDeploymentConfig.User.workstation : "",
            };
            promises.push(this.setupSiteContext(siteDeploymentConfig.Site.Url));
        } else {
            NodeJsomHandlerImpl.authOptions = `Basic ${new Buffer(`${siteDeploymentConfig.User.username}:${siteDeploymentConfig.User.password}`).toString("base64")}`;
        }

        return Promise.all(promises)
            .then(() => {
                return this.loadJsom(siteDeploymentConfig.Site.Url);
            });
    }

    private setupSiteContext(siteUrl: string): Promise<void> {
        // tslint:disable-next-line
        const Agent = siteUrl.indexOf("https") > -1 ? https.Agent : http.Agent;
        const get = siteUrl.indexOf("https") > -1 ? https.get : http.get;
        let reqUrl = Util.JoinAndNormalizeUrl([siteUrl, "_api/web/title"]);
        let parsedUrl = url.parse(reqUrl as string);
        let authValue = NodeJsomHandlerImpl.authOptions;
        if (NodeJsomHandlerImpl.authType === AuthenticationType.NTLM) {
            authValue = NTLM.createType1Message(NodeJsomHandlerImpl.authOptions);
            NodeJsomHandlerImpl.agents[siteUrl] = new Agent({ keepAlive: true, maxSockets: 1, keepAliveMsecs: 100 });
        }
        let options = {
            hostname: parsedUrl.hostname,
            path: parsedUrl.path,
            url: reqUrl,
            method: "GET",
            headers: {
                connection: "keep-alive",
                "Authorization": authValue,
            },
            agent: (NodeJsomHandlerImpl.authType === AuthenticationType.NTLM) ? NodeJsomHandlerImpl.agents[siteUrl] : false,
        };

        return new Promise<void>((resolve, reject) => {
            get(options, firstResponse => {
                firstResponse.on("data", () => undefined);
                firstResponse.on("end", () => {
                    if (firstResponse.statusCode === 401) {
                        let type2msg = NTLM.parseType2Message(firstResponse.headers["www-authenticate"], error => {
                            Util.Reject<void>(reject, "NodeJsom", `JSOM Ntlm initialize error: cannot generate Ntlm type 2 message` + error);
                        });
                        let type3msg = NTLM.createType3Message(type2msg, NodeJsomHandlerImpl.authOptions);
                        options.headers.Authorization = type3msg;
                        get(options, secondResponse => {
                            secondResponse.on("data", () => undefined);
                            secondResponse.on("end", () => {
                                if (secondResponse.statusCode !== 200) {
                                    Util.Reject<void>(reject, "NodeJsom", `JSOM Ntlm initialize error: ${secondResponse.statusCode} after handshake!` + secondResponse.statusMessage);
                                }
                                resolve();
                            });
                        });
                    } else if (firstResponse.statusCode === 200) {
                        resolve();
                    } else {
                        Util.Reject<void>(reject, "NodeJsom", "JSOM Ntlm initialize error: " + firstResponse.statusMessage);
                    }
                });
            }).on("error", error => {
                Util.Reject<void>(reject, "NodeJsom", "JSOM Ntlm initialize error: " + error);
            });
        });
    }

    private loadJsom(siteUrl: string): Promise<void> {
        global.window = global;
        global.window.XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
        let urlObject = url.parse(siteUrl, true, true);
        let relativeUrl = Util.getRelativeUrl(siteUrl);

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
            getElementsByName: (name) => {
                if (name === "__REQUESTDIGEST") {
                    return [global.window.formdigest];
                }
            },
            getElementsByTagName: (name) => {
                return [];
            },
        };

        let scripts = [
            "_layouts/15/init.debug.js",
            "_layouts/15/MicrosoftAjax.js",
            "_layouts/15/sp.core.debug.js",
            "_layouts/15/sp.runtime.debug.js",
            "_layouts/15/sp.debug.js",
            "_layouts/15/sp.publishing.debug.js",
        ];

        return new Promise<void>((resolve, reject) => {
            scripts.reduce((result, currentValue, currentIndex, array) => {
                return result.then((loadedScript: string) => {
                    if (loadedScript) {
                        vm.runInThisContext(loadedScript);
                    }

                    return new Promise((resolveRequest, rejectRequest) => {
                        const get = siteUrl.indexOf("https") > -1 ? https.get : http.get;

                        let reqUrl = Util.JoinAndNormalizeUrl([siteUrl, currentValue]);
                        let parsedUrl = url.parse(reqUrl as string);
                        let authValue = NodeJsomHandlerImpl.authOptions;
                        let options = {
                            hostname: parsedUrl.hostname,
                            path: parsedUrl.path,
                            url: reqUrl,
                            method: "GET",
                            headers: (NodeJsomHandlerImpl.authType === AuthenticationType.NTLM) ?
                                { connection: "keep-alive" } : { "Authorization": authValue },
                            agent: (NodeJsomHandlerImpl.authType === AuthenticationType.NTLM) ? NodeJsomHandlerImpl.agents[siteUrl] : false,
                        };

                        const request = get(options, response => {
                            const body = [];
                            response.on("data", (chunk) => body.push(chunk));
                            response.on("end", () => resolveRequest(body.join("")));
                        });
                        request.on("error", (error) => { Util.Reject<void>(reject, "NodeJsom", "JSOM Ntlm initialize error: " + error); });
                    });
                });
            }, Promise.resolve())
                .then((loadedScript: string) => {
                    if (loadedScript) {
                        vm.runInThisContext(loadedScript);
                    }
                    let context = new SP.ClientContext(siteUrl);
                    let web = context.get_web();
                    context.load(web);
                    context.executeQueryAsync((sender, args) => {
                        resolve();
                    }, (sender, args) => {
                        Util.Reject<void>(reject, "NodeJsom", `Error while initialize JSOM: ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                    });
                });
        });
    }
}

// tslint:disable-next-line
export let NodeJsomHandler: INodeJsomHandler = new NodeJsomHandlerImpl();
