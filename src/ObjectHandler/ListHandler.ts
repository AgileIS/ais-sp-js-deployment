import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";

export class ListHandler implements ISPObjectHandler {
    execute(config: IList, url: string, parent: Promise<ISite>) {
        return new Promise((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config));
        });
    }
}