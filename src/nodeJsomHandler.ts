"use strict";

import { NTLM } from "./ntlm";
import * as http from "http";
import * as https from "https";
import * as url from "url";
import * as vm from "vm";
import { IBasicAuthenticationOptions } from "./interfaces/iBasicAuthenticationOptions";
import { INtlmAuthenticationOptions } from "./interfaces/iNtlmAuthenticationOptions";
import { IPromiseResult } from "./interfaces/iPromiseResult";
import { AuthenticationType } from "./constants/authenticationType";
import { HttpRequestHeader } from "./constants/httpRequestHeader";
import { HttpResponseHeader } from "./constants/httpResponseHeader";
import { HttpMethod } from "./constants/httpMethod";
import { HttpProtocol } from "./constants/httpProtocol";
import { Util } from "./util/util";

declare var global: NodeJS.Global;
declare namespace NodeJS {
    interface Global {  // tslint:disable-line
        window: any;
    }
}

interface INodeJsomHandler {
    initialize(siteUrl: string, webApplicationUrl: string, authenticationType: AuthenticationType,
        authenticationOptions: IBasicAuthenticationOptions | INtlmAuthenticationOptions): Promise<IPromiseResult<void>>;
}

class NodeJsomHandlerImpl implements INodeJsomHandler {
    public static instance: NodeJsomHandlerImpl;
    private static agent: boolean | http.Agent | https.Agent = false;
    private static siteUrl: string;
    private static webApplicationUrl: string;
    private static authenticationType: AuthenticationType;
    private static authenticationOptions: IBasicAuthenticationOptions | INtlmAuthenticationOptions;
    private static federationAuthenticationCookie: string;

    private httpSavedRequest = undefined;
    private httpsSavedRequest = undefined;
    private jsomScripts = [
        "_layouts/15/init.debug.js",
        "_layouts/15/MicrosoftAjax.js",
        "_layouts/15/sp.core.debug.js",
        "_layouts/15/sp.runtime.debug.js",
        "_layouts/15/sp.debug.js",
        "_layouts/15/sp.publishing.debug.js",
    ];

    public static httpRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        if (typeof options !== "string" && !options.protocol) {
            options.protocol = `${HttpProtocol.HTTP}:`;
        }

        return NodeJsomHandlerImpl.instance.httpSavedRequest(NodeJsomHandlerImpl.setRequiredOptions(options), callback);
    }

    public static httpsRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        if (typeof options !== "string" && !options.protocol) {
            options.protocol = `${HttpProtocol.HTTPS}:`;
        }

        return NodeJsomHandlerImpl.instance.httpsSavedRequest(NodeJsomHandlerImpl.setRequiredOptions(options), callback);
    }

    private static setRequiredOptions(options: http.RequestOptions): http.RequestOptions {
        let requestOptions = undefined;
        if (typeof options !== "string") {
            requestOptions = options;
            if (options.headers[HttpRequestHeader.USERAGENT] === "node-XMLHttpRequest" || options.headers[HttpRequestHeader.XREQUESTWITH] === "XMLHttpRequest") {
                requestOptions.agent = NodeJsomHandlerImpl.agent;
                requestOptions.headers = options.headers || {};

                if (!requestOptions.url) {
                    requestOptions.url = Util.JoinAndNormalizeUrl([options.protocol, options.host, options.path]);
                }

                if (NodeJsomHandlerImpl.authenticationType === AuthenticationType.NTLM) {
                    requestOptions.headers[HttpRequestHeader.CONNECTION] = "keep-alive";
                    if (NodeJsomHandlerImpl.federationAuthenticationCookie) {
                        requestOptions.headers[HttpRequestHeader.COOKIE] = NodeJsomHandlerImpl.federationAuthenticationCookie;
                    }
                } else if (NodeJsomHandlerImpl.authenticationType === AuthenticationType.BASIC) {
                    requestOptions.headers[HttpRequestHeader.AUTHORIZATION] = `Basic ${(<IBasicAuthenticationOptions>NodeJsomHandlerImpl.authenticationOptions).encodedUsernamePassword} `;
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

    public initialize(siteUrl: string, webApplicationUrl: string, authenticationType: AuthenticationType,
        authenticationOptions: IBasicAuthenticationOptions | INtlmAuthenticationOptions): Promise<IPromiseResult<void>> {
        // todo: log verbose: Node-Jsom - Start initialize node jsom for site '${siteUrl}'.

        NodeJsomHandlerImpl.siteUrl = siteUrl;
        NodeJsomHandlerImpl.webApplicationUrl = webApplicationUrl;
        NodeJsomHandlerImpl.authenticationType = authenticationType;
        NodeJsomHandlerImpl.authenticationOptions = authenticationOptions;

        let initPromises = new Array<Promise<any>>();
        this.overwriteRequestFunctions();

        if (NodeJsomHandlerImpl.authenticationType === AuthenticationType.NTLM) {
            initPromises.push(this.establishNtmlConnection());
        } else if (NodeJsomHandlerImpl.authenticationType === AuthenticationType.BASIC) {
            let basicAuthenticationOptions = <IBasicAuthenticationOptions>authenticationOptions;
            if (!basicAuthenticationOptions.encodedUsernamePassword) {
                basicAuthenticationOptions.encodedUsernamePassword = `${new Buffer(`${basicAuthenticationOptions.username}:${basicAuthenticationOptions.password}`).toString("base64")} `;
            }
        } else {
            let rejectPromise = new Promise<IPromiseResult<void>>((resolve, reject) => {
                Util.Reject<Promise<any>>(reject, "Node-Jsom", `Error while initialize node-jsom: Unsupported authentication type '${authenticationType}'.`);
            });
            initPromises.push(rejectPromise);
        }

        return Promise.all(initPromises)
            .then(() => {
                this.initializeWindow();
                return this.loadJsomScripts();
            })
            .then(() => {
                return new Promise<IPromiseResult<void>>((resolve, reject) => {
                    Util.Resolve<void>(resolve, "Node-Json", `Initialized node jsom for site '${siteUrl}'.`);
                });
            })
            .catch((error) => {
                return Promise.reject(error);
            });
    }

    private overwriteRequestFunctions(): void {
        // todo: log verbose: Start overwrite http and https request functions.

        this.httpSavedRequest = http.request;
        (<any>http).request = NodeJsomHandlerImpl.httpRequest;

        this.httpsSavedRequest = https.request;
        (<any>https).request = NodeJsomHandlerImpl.httpsRequest;

        // todo: log verbose: Did overwrite http and https request functions.
    }

    private establishNtmlConnection(): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let siteUrl = NodeJsomHandlerImpl.siteUrl;
            let webApplicationUrl = NodeJsomHandlerImpl.webApplicationUrl;

            // todo: log verbose: Start establish ntml connection to ${siteUrl}.

            let authenticationOptions = <INtlmAuthenticationOptions>NodeJsomHandlerImpl.authenticationOptions;
            if (authenticationOptions.domain && authenticationOptions.username && authenticationOptions.password) {
                let get = siteUrl.substr(0, HttpProtocol.HTTPS.length) === HttpProtocol.HTTPS ? https.get : http.get;

                let agentOptions: https.AgentOptions | http.AgentOptions = { keepAlive: true, maxSockets: 1, keepAliveMsecs: 100 };
                NodeJsomHandlerImpl.agent = siteUrl.substr(0, HttpProtocol.HTTPS.length) === HttpProtocol.HTTPS ? new https.Agent(agentOptions) : new http.Agent(agentOptions);

                let reqUrl = Util.JoinAndNormalizeUrl([webApplicationUrl, `_windows/default.aspx?ReturnURl=${encodeURIComponent(siteUrl)}`]);
                let reqUrlObject = url.parse(reqUrl, true, true);

                if (!authenticationOptions.workstation) {
                    authenticationOptions.workstation = "";
                }
                let type1Message = NTLM.createType1Message(authenticationOptions);

                let options = {
                    url: reqUrl,
                    hostname: reqUrlObject.hostname,
                    path: reqUrlObject.path,
                    method: HttpMethod.GET,
                    headers: {
                        [HttpRequestHeader.CONNECTION]: "keep-alive",
                        [HttpRequestHeader.AUTHORIZATION]: type1Message,
                    },
                    agent: NodeJsomHandlerImpl.agent,
                };

                this.sendNtlmType1Message(get, options)
                    .then((type1MsgResult) => {
                        if (type1MsgResult.value) {
                            this.sendNtlmType3Message(get, options, <http.IncomingMessage>type1MsgResult.value)
                                .then((type3MsgResult) => { resolve(type3MsgResult); })
                                .catch((error) => { reject(error); });
                        } else {
                            resolve(<IPromiseResult<void>>type1MsgResult);
                        }
                    }).catch((error) => { reject(error); });
            } else {
                Util.Reject<void>(reject, "Node-Jsom", `Error while establish ntlm connection: Invalid ntlm authentication options.\n` +
                    `Domain, username or password are undefined.\n` +
                    `Authentication options:\n` +
                    `${JSON.stringify(authenticationOptions)}`);
            }
        });
    }

    private sendNtlmType3Message(get: (options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void) => http.ClientRequest,
        requestOptions: http.RequestOptions, type1MsgResponse: http.IncomingMessage): Promise<IPromiseResult<void>> {
        return new Promise<any>((resolve, reject) => {
            let siteUrl = NodeJsomHandlerImpl.siteUrl;
            let type2msg = NTLM.parseType2Message(type1MsgResponse.headers[HttpResponseHeader.WWWAUTHENTICATE.toLocaleLowerCase()], error => {
                Util.Reject<void>(reject, "Node-Jsom", `Error while establish ntlm connection and parsing ntlm type 2 message: ${error}`);
            });

            let type3msg = NTLM.createType3Message(type2msg, NodeJsomHandlerImpl.authenticationOptions);
            requestOptions.headers[HttpRequestHeader.AUTHORIZATION] = type3msg;

            get(requestOptions, type3MsgResponse => {
                type3MsgResponse
                    .on("data", () => undefined)
                    .on("end", () => {
                        if (type3MsgResponse.statusCode === 302) {
                            let redirectionUrl = Util.hasHeader(HttpResponseHeader.LOCATION, type3MsgResponse.headers) ?
                                type3MsgResponse.headers[HttpResponseHeader.LOCATION.toLocaleLowerCase()] : undefined;
                            if (redirectionUrl && redirectionUrl === siteUrl) {
                                let federationAuthenticationCookie = this.getFederationAuthenticationCookie(type3MsgResponse);
                                if (federationAuthenticationCookie) {
                                    NodeJsomHandlerImpl.federationAuthenticationCookie = federationAuthenticationCookie;
                                    Util.Resolve<void>(resolve, "Node-Jsom", "Ntlm connection successfully established with cookie");
                                } else {
                                    Util.Resolve<void>(resolve, "Node-Jsom", "Ntlm connection successfully established without cookie.");
                                }
                            } else {
                                Util.Reject<void>(reject, "Node-Jsom",
                                    `Error while establish ntlm connection and validating ntlm type 3 message response with status code '${type1MsgResponse.statusCode}'.\n` +
                                    `Unexpected redirection url '${redirectionUrl}'.\n` +
                                    `Expected redirection url '${siteUrl}'.`);
                            }
                        } else {
                            Util.Reject<void>(reject, "Node-Jsom",
                                `Error while establish ntlm connection and validating ntlm type 3 message response with status code '${type1MsgResponse.statusCode}.\n` +
                                `Unexpected status code '${type3MsgResponse.statusCode}'.\n` +
                                `Reponse status message: ${type3MsgResponse.statusMessage}`);
                        }
                    });
            }).on("error", error => {
                Util.Reject<void>(reject, "Node-Jsom", `Error while establish ntlm connection and sending ntlm type 3 message: ${error}`);
            });
        });
    }

    private sendNtlmType1Message(get: (options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void) => http.ClientRequest,
        requestOptions: http.RequestOptions): Promise<IPromiseResult<void | http.IncomingMessage>> {
        return new Promise<any>((resolve, reject) => {
            let siteUrl = NodeJsomHandlerImpl.siteUrl;
            get(requestOptions, (type1MsgResponse) => {
                type1MsgResponse
                    .on("data", () => undefined)
                    .on("end", () => {
                        if (type1MsgResponse.statusCode === 401) {
                            let type2Message = Util.hasHeader(HttpResponseHeader.WWWAUTHENTICATE, type1MsgResponse.headers) ?
                                type1MsgResponse.headers[HttpResponseHeader.WWWAUTHENTICATE.toLocaleLowerCase()] : undefined;
                            if (type2Message) {
                                Util.Resolve<http.IncomingMessage>(resolve, "Node-Jsom", "", type1MsgResponse);
                            } else {
                                Util.Reject<void>(reject, "Node-Jsom",
                                    `Error while establish ntlm connection and validating ntlm type 2 message response with status code '${type1MsgResponse.statusCode}.\n` +
                                    `Type 2 message ${HttpResponseHeader.WWWAUTHENTICATE} header is not available.`);
                            }
                        } else if (type1MsgResponse.statusCode === 302) {
                            let redirectionUrl = Util.hasHeader(HttpResponseHeader.LOCATION, type1MsgResponse.headers)
                                ? type1MsgResponse.headers[HttpResponseHeader.LOCATION.toLocaleLowerCase()] : undefined;
                            if (redirectionUrl && redirectionUrl === siteUrl) {
                                let federationAuthenticationCookie = this.getFederationAuthenticationCookie(type1MsgResponse);
                                if (federationAuthenticationCookie) {
                                    NodeJsomHandlerImpl.federationAuthenticationCookie = federationAuthenticationCookie;
                                    Util.Resolve<void>(resolve, "Node-Jsom", "Ntlm connection successfully established with cookie");
                                } else {
                                    Util.Resolve<void>(resolve, "Node-Jsom", "Ntlm connection successfully established without cookie.");
                                }
                            } else {
                                Util.Reject<void>(reject, "Node-Jsom",
                                    `Error while establish ntlm connection and validating ntlm type 2 message response with status code '${type1MsgResponse.statusCode}'.\n` +
                                    `Unexpected redirection url '${redirectionUrl}'.\n` +
                                    `Expected redirection url '${siteUrl}'.`);
                            }
                        } else {
                            Util.Reject<void>(reject, "Node-Jsom",
                                `Error while establish ntlm connection and validating ntlm type 2 message response with status code '${type1MsgResponse.statusCode}.\n` +
                                `Unexpected response status code '${type1MsgResponse.statusCode}'.\n` +
                                `Reponse status message: ${type1MsgResponse.statusMessage} `);
                        }
                    });
            }).on("error", error => {
                Util.Reject<void>(reject, "Node-Jsom", `Error while establish ntlm connection and sending ntlm type 1 message: ${error}`);
            });
        });
    }

    private initializeWindow(): void {
        // todo: log verbose: Node-Jsom - Start initialize global window object with site url '${siteUrl}'.

        let siteUrl = NodeJsomHandlerImpl.siteUrl;
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

        // todo: log verbose: Node-Jsom - Initialized global window object with site url '${siteUrl}'.
    }

    private loadJsomScripts(): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            // todo: log verbose: Node-Jsom - Start loading jsom scripts.
            this.jsomScripts.reduce((dependentPromise, scriptRelativeUrl, index, array) => {
                return dependentPromise.then(() => {
                    return this.loadScript(scriptRelativeUrl);
                });
            }, Promise.resolve())
                .then(() => {
                    return this.CanRequestSite();
                })
                .then(() => {
                    Util.Resolve<void>(resolve, "Node-Jsom", `Loaded jsom scripts.`);
                })
                .catch((error) => {
                    reject(error);
                });
        });
    }

    private loadScript(relativeScriptPath): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let siteUrl = NodeJsomHandlerImpl.siteUrl;
            let get = siteUrl.substr(0, HttpProtocol.HTTPS.length) === HttpProtocol.HTTPS ? https.get : http.get;
            let requestUrl = Util.JoinAndNormalizeUrl([siteUrl, relativeScriptPath]);
            let requestUrlObject = url.parse(requestUrl);
            let headers = {};

            // todo: log verbose: Node-Jsom - Start loading script <requestUrl>.

            if (NodeJsomHandlerImpl.authenticationType === AuthenticationType.NTLM) {
                headers[HttpRequestHeader.CONNECTION] = "keep-alive";
            } else if (NodeJsomHandlerImpl.authenticationType === AuthenticationType.BASIC) {
                headers[HttpRequestHeader.AUTHORIZATION] = `Basic ${(<IBasicAuthenticationOptions>NodeJsomHandlerImpl.authenticationOptions).encodedUsernamePassword}`;
            } else {
                Util.Reject<void>(reject, "Node-Jsom", `Error while loading script: Unsupported authentication type '${NodeJsomHandlerImpl.authenticationType}'.`);
            }

            let options = {
                hostname: requestUrlObject.hostname,
                path: requestUrlObject.path,
                url: requestUrl,
                method: HttpMethod.GET,
                headers: headers,
                agent: NodeJsomHandlerImpl.agent,
            };

            get(options, response => {
                const body = [];
                response
                    .on("data", (chunk) => body.push(chunk))
                    .on("end", () => {
                        let script = body.join("");
                        vm.runInThisContext(script);
                        Util.Resolve<void>(resolve, "Node-Jsom", `Loaded script ${requestUrl}.`, undefined);
                    });
            }).on("error", (error) => {
                Util.Reject<void>(reject, "Node-Jsom", `Error while loading script '${requestUrl}': ${error} `);
            });
        });
    }

    private CanRequestSite(): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            // todo: log verbose: Node-Jsom - Try requesting site with url '${siteUrl}'.

            let siteUrl = NodeJsomHandlerImpl.siteUrl;
            let context = new SP.ClientContext(siteUrl);
            let web = context.get_web();
            context.load(web);
            context.executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, "Node-Jsom", `Successfully requested site with url '${siteUrl}'`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, "Node-Jsom", `Error while try requesting site with url '${siteUrl}'.\n` +
                        `${args.get_message()}\n` +
                        `${args.get_stackTrace()} `);
                });
        });
    }

    private getFederationAuthenticationCookie(responseMessage: http.IncomingMessage) {
        let cookieName = "FedAuth";
        let cookie = "";
        if (responseMessage && responseMessage.headers && Util.hasHeader(HttpResponseHeader.SETCOOKIE, responseMessage.headers)) {
            let setCookieArray = responseMessage.headers[HttpResponseHeader.SETCOOKIE.toLocaleLowerCase()];
            for (let setCookie of setCookieArray) {
                if (setCookie.substring(0, cookieName.length) === cookieName) {
                    cookie = setCookie.substring(0, setCookie.indexOf(";"));
                    break;
                }
            }
        }
        return cookie;
    }
}

// tslint:disable-next-line
export let NodeJsomHandler: INodeJsomHandler = new NodeJsomHandlerImpl();
