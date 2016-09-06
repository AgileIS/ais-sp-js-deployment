import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {IView} from "../interface/Types/IView";
import {IViewField} from "../interface/Types/IViewField";
import * as Utils from "../lib/Util/Util";

export class ViewFieldHandler implements ISPObjectHandler {
    public execute(config: IViewField, url: string, parentConfig: IView) {
        let spWeb = new web.Web(url);
        let element = config;
        let elementAsString = element.toString();
        let parentElement = parentConfig;
        let listId = parentConfig.ParentListId;
        return new Promise((resolve, reject) => {
            spWeb.lists.getById(listId).get().then((result) => {
                if (result) {
                    spWeb.lists.getById(listId).views.filter(`Title eq '${parentElement.Title}'`).select("Id").get().then((result) => {
                        if (result.length === 1) {
                            spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.filter(`Title eq '${elementAsString}'`).get().then((result) => {

                                spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.add(elementAsString).then(() => {
                                    resolve({});
                                    Logger.write(`Viewfield '${elementAsString}' in View ${parentConfig.Title} created`);
                                }).catch(() => {
                                    // Utils.ViewFieldRetry(spWeb, listId, parentConfig.Title, elementAsString, 500).then(() => { }).catch(() => { });
                                    setTimeout(function () {
                                        spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.add(elementAsString).then(() => {
                                            resolve({});
                                            Logger.write(`Viewfield '${elementAsString}' in View ${parentConfig.Title} created`);
                                        }).catch(() => {
                                            setTimeout(function () {
                                                spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.add(elementAsString).then(() => {
                                                    resolve({});
                                                    Logger.write(`Viewfield '${elementAsString}' in View ${parentConfig.Title} created`);
                                                }).catch((error) => {
                                                    Utils.RejectAndLog(error, elementAsString, reject);
                                                });
                                            }, 500);
                                        });
                                    }, 1000);
                                });


                            }).catch((error) => {
                                Utils.RejectAndLog(error, elementAsString, reject);
                            });
                        }
                        else if (result.length === 0) {
                            let error = `Views with Title '${parentElement.Title}' not found`;
                            Utils.RejectAndLog(error, elementAsString, reject);
                        }
                        else {
                            let error = `More than one Views with Title '${parentElement.Title}' found`;
                            Utils.RejectAndLog(error, elementAsString, reject);
                        }
                    }).catch((error) => {
                        Utils.RejectAndLog(error, elementAsString, reject);
                    });
                }
                else {
                    let error = `List with Id '${listId}' does not exist`;
                    Utils.RejectAndLog(error, elementAsString, reject);
                }
            });
        });
    };
}