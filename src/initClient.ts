/**
 * HttpClient Class => ToDo: refactor
 */
import { NTLM } from "./ntlm";

declare var global: any;
declare var require: (path: string) => any;
let nodeFetch = require("node-fetch");
let http = require("http");
let url = require("url");

let _useProxy = false;
let proxy = {
    protocol: "http:",
    hostname: "127.0.0.1",
    port: 8888,
};

let saveRequestObj = http.request;
http.request = function(options){

    if(_useProxy) {
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
            port: options.port,
        });
        options.protocol = proxy.protocol;
        options.hostname = proxy.hostname;
        options.port = proxy.port;
        options.href = null;
        options.host = null;
    }
    return saveRequestObj(options);
};

let ntlmAgent = new http.Agent({ keepAlive: true, maxSockets: 1 });
let spJsom = require("./node-spjsom/index.js");

export class HttpClient {

    public static useProxy = _useProxy;

    public static initAuth(username: string, password: string): void {
        // Fixed missing Header & Request in node
        global.Headers = nodeFetch.Headers;
        global.Request = nodeFetch.Request;
        global.fetch = nodeFetch;

        let userAndDommain = username.split("\\");

        let pnpConfig =  require("@agileis/sp-pnp-js/lib/configuration/pnplibconfig");
        pnpConfig.setRuntimeConfig({
            nodeHttpNtlmClientOptions: {
                username: userAndDommain[1],
                password: password,
                domain: userAndDommain[0],
                workstation: "",
                siteUrl: "",
            },
        });
    }

    public static initJsom(url: string) {
        spJsom.LoadJsom("http://dev2013vm5/sites/demo")
        .then((result) => {
            //auth..
        });
    }
}
