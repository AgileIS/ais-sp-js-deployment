import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {IView} from "../interface/Types/IView";
import {IViewField} from "../interface/Types/IViewField";

export class ViewFieldHandler implements ISPObjectHandler {
    execute(config: IViewField, url: string, parentConfig: IView) {
        let spWeb = new web.Web(url);
        let element = config;
        let elementAsString = element.toString();
        let parentElement = parentConfig;
        let listId = parentConfig.ParentListId;
        return new Promise<IViewField>((resolve, reject) => {
            spWeb.lists.getById(listId).get().then((result) => {
                if (result) {
                    spWeb.lists.getById(listId).views.filter(`Title eq '${parentElement.Title}'`).select("Id").get().then(
                        (result) => {
                            if (result.length === 1) {
                                spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).get().then(
                                    (result) => {
                                        spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.add(elementAsString).then(() => {
                                            resolve(element);
                                        }).catch((error) => {
                                            setTimeout(function () {
                                                spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.add(elementAsString).then(() => {
                                                    resolve(element);
                                                }).catch((error) => {
                                                    setTimeout(function () {
                                                        spWeb.lists.getById(listId).views.getByTitle(parentConfig.Title).fields.add(elementAsString).then(() => {
                                                            resolve(element);
                                                        }).catch((error) => {
                                                            reject(error + " - " + parentElement.Title);
                                                        });
                                                    }, 500);
                                                });
                                            }, 1000);
                                        });
                                    }).catch((error) => {
                                        reject(error + " - " + parentElement.Title);
                                    });
                            }
                            else if (result.length === 0) {
                                let error = `Views with Title '${parentElement.Title}' not found`;
                                reject(error);
                            }
                            else {
                                let error = `More than one Views with Title '${parentElement.Title}' found`;
                                reject(error);
                            }
                        }).catch((error) => {
                            reject(error + " - " + parentElement.Title);
                        });
                }
                else {
                    let error = `List with Id '${listId}' does not exist`;
                    reject(error);
                }
            });
        });
    };
}