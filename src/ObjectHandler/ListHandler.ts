import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import { Logger } from "sp-pnp-js/lib/utils/logging";
import { Web } from "sp-pnp-js/lib/sharepoint/rest/webs";
import { List, Lists } from "sp-pnp-js/lib/sharepoint/rest/lists";
import { Queryable } from "sp-pnp-js/lib/sharepoint/rest/queryable";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IList } from "../interface/Types/IList";
import { ListTemplates } from "../Constants/ListTemplates";
import { ControlOption } from "../Constants/ControlOption";
import { Resolve, Reject } from "../Util/Util";

export class ListHandler implements ISPObjectHandler {
    public execute(listConfig: IList, parentPromise: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parentPromise.then(parentInstance => {
                if (listConfig.TemplateType) {
                    this.ProcessingViewConfig(listConfig, parentInstance).then((listProsssingResult) => { resolve(listProsssingResult); }).catch((error) => { reject(error); });
                }
                else { Reject(reject, `List template type could not be resolved for the list with the internal name ${listConfig.InternalName}`, listConfig.Title); }
            });
        });
    }

    private ProcessingViewConfig(listConfig: IList, parentInstance: Web): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            Logger.write(`Processing ${listConfig.ControlOption === ControlOption.Add || listConfig.ControlOption === undefined ? "Add" : listConfig.ControlOption} list: '${listConfig.Title}'`, Logger.LogLevel.Info);
            parentInstance.lists.filter(`RootFolder/Name eq '${listConfig.InternalName}'`).select("Id").get().then((listRequestResults) => {
                let processingPromise: Promise<List> = undefined;

                if (listRequestResults && listRequestResults.length === 1) {
                    let list = parentInstance.lists.getById(listRequestResults[0].Id);
                    switch (listConfig.ControlOption) {
                        case ControlOption.Update:
                            processingPromise = this.UpdateList(listConfig, parentInstance, list);
                            break;
                        case ControlOption.Delete:
                            processingPromise = this.DeleteList(listConfig, parentInstance, list);
                            break;
                        default:
                            Resolve(reject, `List with the title '${listConfig.Title}' already exists`, listConfig.Title, list);
                            break;
                    }
                } else {
                    switch (listConfig.ControlOption) {
                        case ControlOption.Update:
                        case ControlOption.Delete:
                            Reject(reject, `List with internal name '${listConfig.InternalName}' does not exists`, listConfig.Title);
                            break;
                        default:
                            processingPromise = this.AddList(listConfig, parentInstance);
                            break;
                    }
                }

                if (processingPromise) {
                    processingPromise.then((listProcessingResult) => { resolve(listProcessingResult); }).catch((error) => { reject(error); });
                }
            }).catch((error) => { Reject(reject, `Error while requesting list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
        });
    }

    private AddList(listConfig: IList, parentInstance: Web): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            let properties = this.CreateProperties(listConfig);
            parentInstance.lists.add(listConfig.InternalName, listConfig.Description, listConfig.TemplateType, listConfig.EnableContentTypes, properties).then((listAddResult) => {
                listAddResult.list.update({ Title: listConfig.Title }).then((listUpdateResult) => {
                    Resolve(resolve, `Added list: '${listConfig.Title}'`, listConfig.Title, listUpdateResult.list);
                }).catch((error) => { Reject(reject, `Error while updating list title with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
            }).catch((error) => { Reject(reject, `Error while adding list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
        });
    }

    private UpdateList(listConfig: IList, parentInstance: Web, list: List): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            let properties = this.CreateProperties(listConfig);
            list.update(properties).then((listUpdateResult) => {
                Resolve(resolve, `Updated list: '${listConfig.Title}'`, listConfig.Title, listUpdateResult.list);
            }).catch((error) => { Reject(reject, `Error while updating list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
        });
    }

    private DeleteList(listConfig: IList, parentInstance: Web, list: List): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            list.delete().then(() => {
                Resolve(resolve, `Deleted List: '${listConfig.InternalName}'`, listConfig.Title);
            }).catch((error) => { Reject(reject, `Error while deleting list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
        });
    }

    private CreateProperties(listConfig: IList) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(listConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (listConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.InternalName;
                delete parsedObject.ControlOption;
                delete parsedObject.Field;
                delete parsedObject.EnableContentTypes;
                delete parsedObject.View;
                delete parsedObject.TemplateType;
                break;
            default:
                delete parsedObject.InternalName;
                delete parsedObject.EnableContentTypes;
                delete parsedObject.ControlOption;
                delete parsedObject.Field;
                delete parsedObject.View;
                delete parsedObject.Title;
                delete parsedObject.Description;
                delete parsedObject.TemplateType;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }

}