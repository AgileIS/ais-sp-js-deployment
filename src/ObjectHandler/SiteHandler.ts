import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {ISite} from "../interface/Types/ISite";

export class SiteHandler implements ISPObjectHandler {
    execute(config: ISite, url: string) {
        return new Promise<ISite>((resolve, reject) => {
            let spWeb = new web.Web(url);
            spWeb.lists.get().then((result) => {
                Logger.write("OK - Site is there - go on", 1);
                resolve(config);
                Logger.write("config " + JSON.stringify(config), 0);
            },
                (error) => {
                    Logger.write(error, 0);
                    reject(error);
                });
        });
    };
}