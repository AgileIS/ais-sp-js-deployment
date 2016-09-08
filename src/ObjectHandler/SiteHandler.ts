import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {ISite, ISiteInstance} from "../interface/Types/ISite";
import {RejectAndLog} from "../lib/Util/Util";
import { IInstance } from "../interface/Types/IInstance";

export class SiteHandler implements ISPObjectHandler {
    public execute(config: ISite, url: string, parent: Promise<IInstance>): Promise<ISiteInstance> {
        return new Promise<ISite>((resolve, reject) => {
            parent.then((parentInstance) => {
                let spWeb = new Web(url);
                spWeb.get().then((result) => {
                    Logger.write("web already exists");

                    //TODO: implement logic for Site CRUD
                    resolve(result);
                });
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