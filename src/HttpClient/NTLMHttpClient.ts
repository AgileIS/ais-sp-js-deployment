import {HttpClient, FetchOptions, HttpClientImpl} from "sp-pnp-js/lib/net/HttpClient";
import { Util } from "sp-pnp-js/lib/utils/util";
import { NTLM } from "sp-pnp-js/lib/net/ntlm";
import {Agent} from "http";

class NTLMHttpClientOptions {
    public username: string;
    public password: string;
    public workstation: string;
    public domain: string;
}

class NTLMGlobalAgent {
    private agent: Agent;
    public useGlobal: boolean;

    constructor() {
        this.agent = new Agent({keepAlive: true, maxSockets: 1});
        this.useGlobal = false;
    }

    public getAgent(): Agent {
        return this.agent;
    }
}

class NTLMHttpClient extends HttpClient {
    private impl: HttpClientImpl;
    private tryedToAuth: boolean;
    private keepAliveAgent: Agent;
    private authValue: string;

    constructor() {
        super();
        this.impl = super.getFetchImpl();
        this.tryedToAuth = false;
        this.keepAliveAgent = _agent.useGlobal ? _agent.getAgent() : new Agent({keepAlive: true, maxSockets: 1});
    }

    public fetchRaw(url: string, options?: FetchOptions): Promise<Response> {
        let newHeader = new Headers();
        newHeader.append("Connection", "keep-alive");
        this._mergeHeaders(newHeader, options.headers);

        let extendedOptions = Util.extend(options, { headers: newHeader }, false);
        extendedOptions = Util.extend(extendedOptions, { agent: this.keepAliveAgent }, false);

        let handshake = (ctx): void => {
            extendedOptions.headers.set("Authorization", NTLM.createType1Message(_options));
            this.impl.fetch(url, extendedOptions).then((response) => {
                if (response.status === 401) {
                    let type2msg = NTLM.parseType2Message(response.headers.get("www-authenticate"), (error: Error) => { console.log(error); });
                    this.authValue = NTLM.createType3Message(type2msg, _options);
                    extendedOptions.headers.set("Authorization", this.authValue);
                    retry(ctx);
                }
            }).catch((response) => {
                console.log(response);
                ctx.reject(response);
            });
        };

        let retry = (ctx): void => {
            this.impl.fetch(url, extendedOptions)
                .then((response) => {
                    if (response.status === 401 && !this.tryedToAuth) {
                        handshake(ctx);
                        this.tryedToAuth = true;
                    } else {
                        ctx.resolve(response);
                    }
                }).catch((response) => {
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
            temp.headers.forEach((value: string, name: string) => {
                if (name.toLowerCase() === "accept" && value.toLowerCase() === "application/json") {
                    target.append(name, "application/json;odata=verbose");
                }else {
                    target.append(name, value);
                }
            });
        }
    }
}

let _options = new NTLMHttpClientOptions();
let _agent = new NTLMGlobalAgent();

export let options = _options;
export let agent = _agent;
export let client = NTLMHttpClient;
