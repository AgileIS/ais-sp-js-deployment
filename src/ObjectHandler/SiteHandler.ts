import {ISPObjectHandler} from "../interface/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";

export class SiteHandler implements ISPObjectHandler{
    execute(config: any ){
        return new Promise((resolve, reject) =>{
            Logger.write("config "+ JSON.stringify(config));
        });
    }
}