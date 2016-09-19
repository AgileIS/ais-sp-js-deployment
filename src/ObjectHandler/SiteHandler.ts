import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Queryable } from "@agileis/sp-pnp-js/lib/sharepoint/rest/queryable";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { ISite } from "../interface/Types/ISite";
import { Resolve, Reject } from "../Util/Util";

export class SiteHandler implements ISPObjectHandler {
    public execute(siteConfig: ISite, parentPromise?: Promise<Queryable>): Promise<Web> {
        return new Promise<Web>((resolve, reject) => {
            let web = new Web(siteConfig.Url);
            web.get()
                .then((result) => {
                    // todo: implement logic for Site CRUD
                    Resolve(resolve, `Web '${siteConfig.Url}' already exists`, siteConfig.Url, web);
                })
                .catch((error) => { Reject(reject, `Error while requesting web with the url '${siteConfig.Url}': ` + error, siteConfig.Url); });
        });
    };
}
