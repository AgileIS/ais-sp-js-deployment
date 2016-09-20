/// <reference path="../typings/index.d.ts" />

import * as http from "http";
import * as https from "https";
import * as url from "url";

interface NodeHttpProxy {
    url: url.Url;
    isActive: boolean;
    activate();
    deactivate();
}

class NodeHttpProxyImpl implements NodeHttpProxy {
    public static instance: NodeHttpProxyImpl;

    public url: url.Url;

    private _httpSavedRequest = undefined;
    private _httpsSavedRequest = undefined;
    private _isActive: boolean = false;

    public static httpRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        NodeHttpProxyImpl.setupProxy(options);
        return NodeHttpProxyImpl.instance._httpSavedRequest(options, callback);
    }

    public static httpsRequest(options: http.RequestOptions, callback?: (res: http.IncomingMessage) => void): http.ClientRequest {
        NodeHttpProxyImpl.setupProxy(options);
        return NodeHttpProxyImpl.instance._httpsSavedRequest(options, callback);
    }

    public static setupProxy(options: http.RequestOptions): void {
        if (NodeHttpProxyImpl.instance.isActive) {
            let requestOptions = undefined;

            if (typeof options === "string") {
                url.parse(options as string, true);
            } else {
                requestOptions = options;
            }

            if (!options.host && !options.hostname) {
                throw new Error("host or hostname must have value.");
            }

            options.path = url.format(requestOptions.url);
            options.headers = options.headers || {};

            requestOptions.headers.Host = requestOptions.host || url.format({
                hostname: requestOptions.hostname,
                port: requestOptions.port,
            });

            requestOptions.protocol = NodeHttpProxyImpl.instance.url.protocol;
            requestOptions.hostname = NodeHttpProxyImpl.instance.url.hostname;
            requestOptions.port = Number(NodeHttpProxyImpl.instance.url.port);
            requestOptions.href = null;
            requestOptions.host = null;
        }
    }

    public constructor() {
        NodeHttpProxyImpl.instance = this;
    }

    public get isActive(): boolean {
        return this._isActive;
    }

    public activate() {
        if (!this._isActive) {
            this._httpSavedRequest = http.request;
            http.request = NodeHttpProxyImpl.httpRequest;

            this._httpsSavedRequest = https.request;
            https.request = NodeHttpProxyImpl.httpsRequest;

            this._isActive = true;
        }
    }

    public deactivate() {
        if (this._isActive) {
            http.request = this._httpSavedRequest;
            this._httpSavedRequest = () => { ; };

            https.request = this._httpsSavedRequest;
            this._httpsSavedRequest = () => { ; };

            this._isActive = false;
        }
    }
}


export let NodeHttpProxy: NodeHttpProxy = new NodeHttpProxyImpl();




