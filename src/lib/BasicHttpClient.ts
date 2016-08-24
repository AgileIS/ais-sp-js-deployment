import {HttpClient, FetchOptions} from "sp-pnp-js/lib/net/HttpClient"
import {DigestCache} from "sp-pnp-js/lib/net/DigestCache"
import * as Util from "sp-pnp-js/lib/utils/util"

class BasicHttpClientOptions {
    username : string;
    password: string;
}

class BasicHttpClient extends HttpClient {
    private authKey: string = "Authorization";
    private authValue: string;

    constructor(){
        super();
        this.authValue = `Basic ${btoa(`${_options.username}:${_options.password}`)}`
    }

    fetchRaw(url: string, options?: FetchOptions): Promise<Response>{
        let newHeader = new Headers();
        newHeader.append(this.authKey, this.authValue);
        this._mergeHeaders(newHeader, options.headers);
        
        let extendedOptions = Util.Util.extend(options, { headers: newHeader }, false);
        return super.fetchRaw(url, extendedOptions);
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
            });
        }
    }
}

let _options = new BasicHttpClientOptions();

export let options = _options;
export let client = BasicHttpClient;


