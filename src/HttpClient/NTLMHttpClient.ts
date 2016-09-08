import {HttpClient, FetchOptions, HttpClientImpl} from "sp-pnp-js/lib/net/HttpClient"
import {DigestCache} from "sp-pnp-js/lib/net/DigestCache"
import { Util } from "sp-pnp-js/lib/utils/util"
import * as NTLM from "./ntlm" 
import {Agent, AgentOptions} from "http";

class NTLMHttpClientOptions {
    username : string;
    password: string;
    workstation: string;
    domain: string;
}

class NTLMHttpClient extends HttpClient {
    private authKey: string = "Authorization";
    private authValue: string;
    private agent: Agent;
    private impl: HttpClientImpl;

    constructor(){
        super();
        this.impl = super.getFetchImpl();
        this.agent = new Agent({ keepAlive: true });
        //ToDo Handshake
    }

    fetchRaw(url: string, options?: FetchOptions): Promise<Response>{
        if(!this.authValue) this.authValue = NTLM.createType1Message(_options);
        let newHeader = new Headers();
        newHeader.append(this.authKey, this.authValue);
        this._mergeHeaders(newHeader, options.headers);
        
        let extendedOptions = Util.extend(options, { headers: newHeader }, false);
        extendedOptions = Util.extend(extendedOptions, { agent: this.agent }, false);

        let retry = (ctx): void => {
            this.impl.fetch(url, extendedOptions).then((response) => {
                if (response.status == 401) {
                    var type2msg = NTLM.parseType2Message(response.headers.get("www-authenticate"), (error: Error) => {});
                    this.authValue = NTLM.createType3Message(type2msg, _options);
                    extendedOptions.headers.set(this.authKey, this.authValue);
                    this.impl.fetch(url, extendedOptions).then((response) => ctx.resolve(response)).catch((response) => {
                        // grab our current delay
                        let delay = ctx.delay;
                        // Check if request was throttled - http status code 429 
                        // Check is request failed due to server unavailable - http status code 503 
                        if (response.status !== 429 && response.status !== 503) {
                            ctx.reject(response);
                        }
                        // Increment our counters.
                        ctx.delay *= 2;
                        ctx.attempts++;
                        // If we have exceeded the retry count, reject.
                        if (ctx.retryCount <= ctx.attempts) {
                            ctx.reject(response);
                        }
                        // Set our retry timeout for {delay} milliseconds.
                        setTimeout(Util.getCtxCallback(this, retry, ctx), delay);
                    });     
                }
            }).catch((response) => {
                console.log(response);
                ctx.reject(response);
            });
        };


        return new Promise((resolve, reject) => {

            let retryContext = {
                attempts: 0,
                delay: 100,
                reject: reject,
                resolve: resolve,
                retryCount: 7,
            };

            retry.call(this, retryContext);
        });
    }

    private _mergeHeaders(target: Headers, source: any): void {
        if (typeof source !== "undefined" && source !== null) {
            let temp = <any>new Request("", { headers: source });
            temp.headers.forEach((value :string, name: string) => {
                if(name.toLowerCase() == "accept" && value.toLowerCase() == "application/json"){
                    target.append(name, "application/json;odata=verbose");   
                }else{
                    target.append(name, value);   
                }
                target.append("Connection", "keep-alive"); 
            });
        }
    }
}

let _options = new NTLMHttpClientOptions();

export let options = _options;
export let client = NTLMHttpClient;


/*var options = {
    url: "http://dev2013vm5/sites/insidervz",
    username: "tuser",
    password: "Start@123",
    workstation: "",
    domain: "dev"
};
var ntlmToken = "";
var request = require('request');
var HttpAgent = require('agentkeepalive');
var keepaliveAgent = new HttpAgent();


// Set the headers
var headers = {
    "Authorization": ntlm.createType1Message(options),
    "Accept": "application/json;odata=verbose",
    "Connection": "keep-alive"
}

// Configure the request
var options2 = {
    url: "http://dev2013vm5/sites/insidervz/_api/web",
    method: 'GET',
    headers: headers,
    agent: keepaliveAgent
}

// Start the request
request(options2, function (error, response, body) {
    if (!error && response.statusCode == 401) {
        var type2msg = ntlm.parseType2Message(response.headers['www-authenticate']);
        ntlmToken = ntlm.createType3Message(type2msg, options);

        headers["Authorization"] = ntlmToken;
        options2.headers = headers;

        request(options2, function (error, response, body) {
            if (!error && response.statusCode == 200) {
                console.log(body);
            } else {
                console.log(response.toJSON());
            } 
        });
    }
})*/