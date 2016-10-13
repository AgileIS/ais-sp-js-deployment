/// <reference path="../typings/index.d.ts" />

// import * as http from "http";
// import * as https from "https";
import http = require("http");
import https = require("https");
import * as url from "url";

interface INodeHttpProxy {
    url: url.Url;
    isActive: boolean;
    activate();
    deactivate();
}

class NodeHttpProxyImpl implements INodeHttpProxy {
    public static instance: NodeHttpProxyImpl;

    public url: url.Url;

    private httpSavedRequest = undefined;
    private httpsSavedRequest = undefined;
    private active: boolean = false;

    public static httpRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        return NodeHttpProxyImpl.instance.httpSavedRequest(NodeHttpProxyImpl.setupProxy(options), callback);
    }

    public static httpsRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        return NodeHttpProxyImpl.instance.httpsSavedRequest(NodeHttpProxyImpl.setupProxy(options), callback);
    }

    public static setupProxy(options: http.RequestOptions): http.RequestOptions {
        if (NodeHttpProxyImpl.instance.isActive) {
            let requestOptions = undefined;

            if (typeof options === "string") {
                requestOptions = url.parse(options as string, true);
                requestOptions.url = options;
            } else {
                requestOptions = options;
            }

            if (!requestOptions.host && !requestOptions.hostname) {
                throw new Error("host or hostname must have value.");
            }

            requestOptions.path = url.format(requestOptions.url);
            requestOptions.headers = options.headers || {};

            requestOptions.headers.Host = requestOptions.host || requestOptions.hostname;

            requestOptions.protocol = NodeHttpProxyImpl.instance.url.protocol;
            requestOptions.hostname = NodeHttpProxyImpl.instance.url.hostname;
            requestOptions.port = Number(NodeHttpProxyImpl.instance.url.port);
            requestOptions.href = undefined;
            requestOptions.host = undefined;
            return requestOptions;
        }
        return options;
    }

    public constructor() {
        NodeHttpProxyImpl.instance = this;
    }

    public get isActive(): boolean {
        return this.active;
    }

    public activate() {
        if (!this.active) {
            this.httpSavedRequest = http.request;
            http.request = NodeHttpProxyImpl.httpRequest;

            this.httpsSavedRequest = https.request;
            https.request = NodeHttpProxyImpl.httpsRequest;

            this.active = true;
        }
    }

    public deactivate() {
        if (this.active) {
            http.request = this.httpSavedRequest;
            this.httpSavedRequest = () => { ; };

            https.request = this.httpsSavedRequest;
            this.httpsSavedRequest = () => { ; };

            this.active = false;
        }
    }
}

// tslint:disable-next-line
export let NodeHttpProxy: INodeHttpProxy = new NodeHttpProxyImpl();
