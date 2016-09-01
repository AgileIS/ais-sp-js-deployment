import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {IView} from "../interface/Types/IView";
import {IList} from "../interface/Types/IList";


export class ViewHandler implements ISPObjectHandler {
    execute(config: IView, url: string, parentConfig: IList) {
        let spWeb = new web.Web(url);
        let element = config;
        let listName = parentConfig.InternalName;
        return new Promise<IView>((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    let propertyHash = createTypedHashfromProperties(element);
                    spWeb.lists.getById(listId).views.filter(`Title eq '${element.Title}'`).select("Id").get().then(
                        (result) => {
                            if (result.length === 1) {
                                resolve(config);
                                Logger.write(`View with Title '${element.Title}' already exists`);
                            }
                            else if (result.length === 0) {
                                spWeb.lists.getById(listId).views.add(element.Title, element.PersonalView, propertyHash).then(
                                    (result) => {
                                        result.view.fields.removeAll().then(
                                            () => {
                                                let promiseArray = [];
                                                for (let viewField of element.ViewFields) {
                                                    promiseArray.push(spWeb.lists.getById(listId).views.getByTitle(element.Title).fields.add(viewField));  // Timing Problem
                                                }
                                                Promise.all(promiseArray).then(
                                                    () => {
                                                        resolve(config);
                                                    },
                                                    (error) => {
                                                        reject(error + " - " + element.Title);
                                                    }
                                                );
                                            },
                                            (error) => {
                                                reject(error + " - " + element.Title);
                                            }
                                        );
                                    },
                                    (error) => {
                                        reject(error + " - " + element.Title);
                                    }
                                );
                            }
                            else {
                                let error = `More than one Views wit Title '${element.Title}' found`;
                                reject(error);
                            }
                        },
                        (error) => {
                            reject(error + " - " + element.Title);
                        }
                    );
                }
                else {
                    let error = `List with Internal Name '${listName}' does not exist`;
                    reject(error);
                }
            });
        });
    };
}

function createTypedHashfromProperties(pElement: IView) {
    let element = pElement;
    let stringifiedObject: string;
    stringifiedObject = JSON.stringify(element);
    let parsedObject = JSON.parse(stringifiedObject);
    delete parsedObject.ControlOption;
    delete parsedObject.Title;
    delete parsedObject.PersonalView;
    delete parsedObject.ViewFields;
    stringifiedObject = JSON.stringify(parsedObject);
    return JSON.parse(stringifiedObject);
}