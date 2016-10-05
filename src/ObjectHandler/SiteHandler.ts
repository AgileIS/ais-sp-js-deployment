import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISite } from "../Interfaces/Types/ISite";
import * as PnP from "@agileis/sp-pnp-js";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { Util } from "../Util/Util";

export class SiteHandler implements ISPObjectHandler {
    public execute(siteConfig: ISite, parentPromise: Promise<IPromiseResult<void>>): Promise<IPromiseResult<Web>> {
        return new Promise<IPromiseResult<Web>>((resolve, reject) => {
            if (siteConfig && siteConfig.Url) {
                PnP.sp.web.get()
                    .then((result) => {
                        Util.Resolve<Web>(resolve, siteConfig.Url, `Web '${siteConfig.Url}' already exists.`, PnP.sp.web);
                    })
                    .catch((error) => { Util.Reject<void>(reject, siteConfig.Url, `Error while requesting web with the url '${siteConfig.Url}': ` + Util.getErrorMessage(error)); });
            } else {
                Util.Reject<void>(reject, siteConfig.Url, `Error while processing site with the url '${siteConfig.Url}': site url is undefined.`);
            }
        });
    };
}
