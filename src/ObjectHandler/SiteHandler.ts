import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {ISite} from "../interface/Types/ISite";

export class SiteHandler implements ISPObjectHandler {
    execute(config: any, url: string, parent: Promise<any>) {
        let parentPromise = parent;
        let promise: Promise<ISite>;
        return new Promise<ISite>((resolve, reject) => {
            parentPromise.then(() => {
                let spWeb = new web.Web(url);
                if (spWeb) {
                    resolve();
                    Logger.write("config " + JSON.stringify(config));
                }
            },
                () => {
                    let error = "Parent Promise not resolved";
                    reject(error);
                    Logger.write(error);
                }
            )
        })
    };
}