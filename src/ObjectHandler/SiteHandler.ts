import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {ISite} from "../interface/Types/ISite";
import {RejectAndLog} from "../Util/Util";
import { Queryable } from "sp-pnp-js/lib/sharepoint/rest/queryable";


export class SiteHandler implements ISPObjectHandler {
    public execute(config: ISite, parent?: Promise<Queryable>): Promise<Web> {
        return new Promise<Web>((resolve, reject) => {
                let spWeb = new Web(config.Url);
                spWeb.get().then((result) => {
                    Logger.write("web already exists");
                    //TODO: implement logic for Site CRUD
                    resolve(spWeb);
            });

            // spWeb.lists.get().then((result) => {
            //     resolve(config);
            //     Logger.write("config " + JSON.stringify(config), 0);
            // }).catch((error) => {
            //     RejectAndLog(error, "Site", reject);
            // });
        });
    };
}