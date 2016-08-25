import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {ISite} from "../interface/Types/ISite";

export class SiteHandler implements ISPObjectHandler {
    execute(config: ISite, url: string) {
        return new Promise<ISite>((resolve, reject) => {
           
                 let spWeb = new web.Web(url);
                if (spWeb) {
                    Logger.write("OK - Site is there - go on");
                    resolve(config);
                    Logger.write("config " + JSON.stringify(config));
                }
                else {
                    reject("Site not found");
                }
           
        }
        );
    };
}