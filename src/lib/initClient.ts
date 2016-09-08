/**
 * HttpClient Class
 */
export class HttpClient {
        public static initAuth(username: string, password: string) {
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

        let httpBasicClient = require("./BasicHttpClient");

        let userAndDommain = username.split("\\");

        httpBasicClient.options.username = userAndDommain[0] + "\\" + userAndDommain[1];
        httpBasicClient.options.password = password;

        let httpClient = require("sp-pnp-js/lib/net/HttpClient");
        httpClient.HttpClient = httpBasicClient.client;
    }
}

