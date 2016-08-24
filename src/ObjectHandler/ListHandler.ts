import {ISPObjectHandler} from "../interface/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IListInstance} from "sp-pnp-js/lib/sharepoint/provisioning/schema/ilistinstance";

export class ListHandler implements ISPObjectHandler{
    execute(config: IListInstance ){
        return new Promise((resolve, reject) =>{
            Logger.write("config "+ JSON.stringify(config));
        });
    }
}