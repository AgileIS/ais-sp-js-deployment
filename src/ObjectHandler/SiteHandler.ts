import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {ISite} from "../interface/Types/ISite";
import {RejectAndLog} from "../lib/Util/Util";

export class SiteHandler implements ISPObjectHandler {
    public execute(config: ISite, url: string, parentConfig: any) {
        return new Promise<ISite>((resolve, reject) => {
            let spWeb = new Web(url);
            spWeb.lists.get().then((result) => {
                resolve(config);
                Logger.write("config " + JSON.stringify(config), 0);
            }).catch((error) => {
                RejectAndLog(error, "Site", reject);
            });
        });
    };
}