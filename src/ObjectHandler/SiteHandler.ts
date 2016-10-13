import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISite } from "../Interfaces/Types/ISite";
import * as PnP from "@agileis/sp-pnp-js";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { Util } from "../Util/Util";
import { ControlOption } from "../Constants/ControlOption";

export class SiteHandler implements ISPObjectHandler {
    public execute(siteConfig: ISite, parentPromise: Promise<IPromiseResult<void>>): Promise<IPromiseResult<Web>> {
        return new Promise<IPromiseResult<Web>>((resolve, reject) => {
            if (siteConfig && siteConfig.Url) {
                PnP.sp.web.get()
                    .then((result) => {
                        switch (siteConfig.ControlOption) {
                            case ControlOption.UPDATE:
                                this.updateSiteProperties(siteConfig, PnP.sp.web)
                                    .then((siteProcessingResult) => { resolve(siteProcessingResult); })
                                    .catch((error) => { reject(error); });
                                break;
                            case ControlOption.DELETE:
                                Util.Reject<void>(reject, siteConfig.Url, `Error delete a site is not supported`);
                                break;
                            default:
                                Util.Resolve<Web>(resolve, siteConfig.Url, `Web '${siteConfig.Url}' already exists.`, PnP.sp.web);
                        }
                    })
                    .catch((error) => { Util.Reject<void>(reject, siteConfig.Url, `Error while requesting web with the url '${siteConfig.Url}': ` + Util.getErrorMessage(error)); });
            } else {
                Util.Reject<void>(reject, "Unknown site", `Error while processing site: site url is undefined.`);
            }
        });
    };

    private updateSiteProperties(siteConfig: ISite, site: Web): Promise<IPromiseResult<Web>> {
        return new Promise<IPromiseResult<Web>>((resolve, reject) => {
            let properties = this.createProperties(siteConfig);
            site.update(properties)
                .then((siteUpdateResult) => {
                    if (siteConfig.PropertyBagEntries) {
                        this.updatePropertyBag(siteConfig, siteUpdateResult.web)
                            .then(() => {
                                Util.Resolve<Web>(resolve, siteConfig.Title, `Updated site properties: '${siteConfig.Title}' and added PropertyBagEntries.`, siteUpdateResult.web);
                            })
                            .catch((error) => {
                                Util.Reject<void>(reject, siteConfig.Title,
                                    `Error while adding PropertyBagEntries in the site with the title '${siteConfig.Title}': ` + Util.getErrorMessage(error));
                            });
                    } else {
                        Util.Resolve<Web>(resolve, siteConfig.Title, `Updated site properties: '${siteConfig.Title}'.`, siteUpdateResult.web);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, siteConfig.Title, `Error while updating site with the title '${siteConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private updatePropertyBag(siteConfig: ISite, site: Web): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let context = SP.ClientContext.get_current();
            let web = context.get_web();
            let propBag = web.get_allProperties();

            for (let prop of siteConfig.PropertyBagEntries) {
                propBag.set_item(prop.Title, prop.Value);
            }
            web.update();
            context.executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, siteConfig.Title, `Updated property bag entries in site with the title '${siteConfig.Title}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, siteConfig.Title, `Error while updating property bag in the site with the title '${siteConfig.Title}': `
                            + `${Util.getErrorMessageFromQuery(args.get_message(),args.get_stackTrace())}`);
                });
        });
    }

    private createProperties(viewConfig: ISite) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(viewConfig);
        let parsedObject: ISite = JSON.parse(stringifiedObject);

        delete parsedObject.ControlOption;
        delete parsedObject.Url;
        delete parsedObject.ContentTypes;
        delete parsedObject.Lists;
        delete parsedObject.Files;
        delete parsedObject.Fields;
        delete parsedObject.Navigation;
        delete parsedObject.CustomActions;
        delete parsedObject.ComposedLook;
        delete parsedObject.PropertyBagEntries;
        delete parsedObject.Parameters;
        delete parsedObject.Features;
        delete parsedObject.WebSettings;
        delete parsedObject.InheritPermissions;
        delete parsedObject.Language;
        delete parsedObject.Template;
        delete parsedObject.WebSettings;
        delete parsedObject.LayoutsHive;
        delete parsedObject.Identifier;

        switch (viewConfig.ControlOption) {
            case ControlOption.UPDATE:
                break;
            default:
                delete parsedObject.Title;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}
