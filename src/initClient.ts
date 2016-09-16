/**
 * HttpClient Class
 */
declare var global: any;
declare var require: (path: string) => any;
let nodeFetch = require("node-fetch");
let http = require("http");
let url = require("url");

let saveRequestObj = http.request;
let proxy = {
    protocol: "http:",
    hostname: "127.0.0.1",
    port: 8888,
};

export class HttpClient {
    public static initAuth(username: string, password: string) {
        // Fixed missing Header & Request in node
        global.Headers = nodeFetch.Headers;
        global.Request = nodeFetch.Request;
        global.fetch = nodeFetch;

        let userAndDommain = username.split("\\");

        let pnpConfig =  require("sp-pnp-js/lib/configuration/pnplibconfig");
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

    public static useProxy() {
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
                port: options.port,
            });
            options.protocol = proxy.protocol;
            options.hostname = proxy.hostname;
            options.port = proxy.port;
            options.href = null;
            options.host = null;
            return saveRequestObj(options);
        };
    }
}
