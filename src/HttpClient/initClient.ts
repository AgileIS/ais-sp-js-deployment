/**
 * HttpClient Class
 */
export class HttpClient {
        public static initAuth(username: string, password: string) {
        // Fixed missing Header & Request in node
        let fetch = require("node-fetch");
        global["Headers"] = fetch.Headers;
        global["Request"] = fetch.Request;
        global["fetch"] = fetch;

        let userAndDommain = username.split("\\");

        let httpNTLMClient = require("./NTLMHttpClient");
        httpNTLMClient.options.username = userAndDommain[1];
        httpNTLMClient.options.password = password;
        httpNTLMClient.options.domain = userAndDommain[0];
        httpNTLMClient.options.workstation = "";

        let httpBasicClient = require("./BasicHttpClient");
        httpBasicClient.options.username = userAndDommain[0] + "\\" + userAndDommain[1];
        httpBasicClient.options.password = password;

        let httpClient = require("sp-pnp-js/lib/net/HttpClient");
        httpClient.HttpClient = httpBasicClient.client;
        // httpClient.HttpClient = httpNTLMClient.client;
    }
}

