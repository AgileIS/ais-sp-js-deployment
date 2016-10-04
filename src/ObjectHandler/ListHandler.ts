import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { List} from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { ISPObjectHandler } from "../interfaces/ObjectHandler/ispobjecthandler";
import { IList } from "../interfaces/Types/IList";
import { IContentTypeBinding } from "../interfaces/Types/IContentTypeBinding";
import { ListTemplates } from "../Constants/ListTemplates";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { ControlOption } from "../Constants/ControlOption";
import { Util } from "../Util/Util";

export class ListHandler implements ISPObjectHandler {
    public execute(listConfig: IList, parentPromise: Promise<IPromiseResult<Web>>): Promise<IPromiseResult<void | List>> {
        return new Promise<IPromiseResult<void | List>>((resolve, reject) => {
            parentPromise.then(promiseResult => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, listConfig.InternalName,
                        `List handler parent promise value result is null or undefined for the list with the internal name '${listConfig.InternalName}'!`);
                } else {
                    let web = promiseResult.value;
                    this.processingListConfig(listConfig, web)
                        .then((listProsssingResult) => { resolve(listProsssingResult); })
                        .catch((error) => {
                            Util.Retry(error, listConfig.InternalName,
                                () => {
                                    return this.processingListConfig(listConfig, web);
                                });
                        });
                }
            });
        });
    }

    private processingListConfig(listConfig: IList, web: Web): Promise<IPromiseResult<void | List>> {
        return new Promise<IPromiseResult<void | List>>((resolve, reject) => {
            let processingText = listConfig.ControlOption === ControlOption.Add || listConfig.ControlOption === undefined || listConfig.ControlOption === ""
                ? "Add" : listConfig.ControlOption;
            Logger.write(`Processing ${processingText} list: '${listConfig.InternalName}'.`, Logger.LogLevel.Info);

            web.lists.filter(`RootFolder/Name eq '${listConfig.InternalName}'`).select("Id").get()
                .then((listRequestResults) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | List>> = undefined;
                    if (listRequestResults && listRequestResults.length === 1) {
                        Logger.write(`Found List with title: '${listConfig.Title}'`);
                        let list = web.lists.getById(listRequestResults[0].Id);
                        switch (listConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateList(listConfig, list);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteList(listConfig, list);
                                break;
                            default:
                                Util.Resolve<List>(resolve, listConfig.InternalName, `List with internal name '${listConfig.InternalName}'` +
                                    ` does not have to be added, because it already exists.`, list);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (listConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Resolve<void>(resolve, listConfig.InternalName, `List with internal name '${listConfig.InternalName}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.Update:
                                listConfig.ControlOption = ControlOption.Add;
                            default:
                                processingPromise = this.addList(listConfig, web);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((listProcessingResult) => { resolve(listProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write("List handler processing promise is undefined!", Logger.LogLevel.Error);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, listConfig.InternalName, `Error while requesting list with the internal name '${listConfig.InternalName}': ` + error); });
        });
    }

    private addList(listConfig: IList, web: Web): Promise<IPromiseResult<List>> {
        return new Promise<IPromiseResult<List>>((resolve, reject) => {
            if (listConfig.TemplateType) {
                let properties = this.createProperties(listConfig);
                web.lists.add(listConfig.InternalName, listConfig.Description, ListTemplates[(listConfig.TemplateType as string)], listConfig.EnableContentTypes, properties)
                    .then((listAddResult) => {
                        listConfig.ControlOption = ControlOption.Update;
                        this.updateList(listConfig, listAddResult.list)
                            .then((listUpdateResult) => { Util.Resolve<List>(resolve, listConfig.InternalName, `Added list: '${listConfig.InternalName}'.`, listUpdateResult.value); })
                            .catch((error) => {
                                Util.Reject<void>(reject, listConfig.InternalName,
                                    `Error while adding and updating list title with the internal name '${listConfig.InternalName}': ` + error);
                            });
                    })
                    .catch((error) => { Util.Reject<void>(reject, listConfig.InternalName, `Error while adding list with the internal name '${listConfig.InternalName}': ` + error); });
            } else {
                Util.Reject<void>(reject, listConfig.InternalName, `List template type could not be resolved for the list with the internal name ${listConfig.InternalName}`);
            }
        });
    }

    private updateList(listConfig: IList, list: List): Promise<IPromiseResult<List>> {
        return new Promise<IPromiseResult<List>>((resolve, reject) => {
            let properties = this.createProperties(listConfig);
            list.update(properties)
                .then((listUpdateResult) => {
                    if (listConfig.ContentTypeBindings && listConfig.ContentTypeBindings.length > 0) {
                        this.updateListContentTypes(listConfig, list)
                            .then((contentTypesUpdateResult) => { Util.Resolve<List>(resolve, listConfig.InternalName, `Updated list: '${listConfig.InternalName}'.`, listUpdateResult.list); })
                            .catch((error) => {
                                Util.Reject<void>(reject, listConfig.InternalName,
                                    `Error while updating list title with the internal name '${listConfig.InternalName}': ` + error);
                            });
                    } else {
                        Util.Resolve<List>(resolve, listConfig.InternalName, `Updated list: '${listConfig.InternalName}'.`, listUpdateResult.list);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, listConfig.InternalName, `Error while updating list with the internal name '${listConfig.InternalName}': ` + error); });
        });
    }

    private updateListContentTypes(listConfig: IList, list: List): Promise<IPromiseResult<List>> {
        return new Promise<IPromiseResult<List>>((resolve, reject) => {
            let processingPromises: Array<Promise<any>> = [];
            listConfig.ContentTypeBindings.forEach(element => {
                if (element.Delete) {
                    processingPromises.push(this.deleteListContentType(element, listConfig, list));
                } else {
                    processingPromises.push(list.contentTypes.addById(element.ContentTypeId));
                }
            });
            Promise.all(processingPromises)
                .then(() => { Util.Resolve<List>(resolve, listConfig.InternalName, `Updated List ContentTypes: '${listConfig.InternalName}'.`); })
                .catch((error) => { Util.Reject<void>(reject, listConfig.InternalName, `Error while updating list contenttypes, list internal name '${listConfig.InternalName}': ` + error); });
        });
    }

    private deleteListContentType(ctBinding: IContentTypeBinding, listConfig: IList, list: List): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            if (ctBinding.ContentTypeName) {
                list.contentTypes.filter(`Name+eq+'${ctBinding.ContentTypeName}'`).select("Id").get()
                    .then((ctRequestResults) => {
                        if (ctRequestResults && ctRequestResults.length === 1) {
                            list.contentTypes.getById(ctRequestResults[0].Id.StringValue).delete()
                                .then(() => {
                                    Util.Resolve<void>(resolve, listConfig.InternalName,
                                        `Deleted List ContentType: '${ctBinding.ContentTypeName}'from list: '${listConfig.InternalName}'.`);
                                })
                                .catch((error) => {
                                    Util.Reject<void>(reject, listConfig.InternalName,
                                        `Error while deleting list contenttype '${ctBinding.ContentTypeName}', list internal name '${listConfig.InternalName}': ` + error);
                                });
                        } else {
                            Util.Reject<void>(reject, ctBinding.ContentTypeName,
                                `Error while deleting list contenttype '${ctBinding.ContentTypeName}', list internal name '${listConfig.InternalName}': `);
                        }
                    })
                    .catch((error) => {
                        Util.Reject<void>(reject, listConfig.InternalName,
                            `Error while requesting list contenttypes, list internal name '${listConfig.InternalName}': ` + error);
                    });
            } else {
                list.contentTypes.getById(ctBinding.ContentTypeId).delete()
                    .then(() => {
                        Util.Resolve<void>(resolve, listConfig.InternalName,
                            `Deleted List ContentType: '${ctBinding.ContentTypeId}'from list: '${listConfig.InternalName}'.`);
                    })
                    .catch((error) => {
                        Util.Reject<void>(reject, listConfig.InternalName,
                            `Error while deleting list contenttype '${ctBinding.ContentTypeId}', list internal name '${listConfig.InternalName}': ` + error);
                    });
            }
        });
    };

    private deleteList(listConfig: IList, list: List): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            list.delete()
                .then(() => { Util.Resolve<void>(resolve, listConfig.InternalName, `Deleted List: '${listConfig.InternalName}'.`); })
                .catch((error) => { Util.Reject<void>(reject, listConfig.InternalName, `Error while deleting list with the internal name '${listConfig.InternalName}': ` + error); });
        });
    }

    private createProperties(listConfig: IList) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(listConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (listConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.InternalName;
                delete parsedObject.ControlOption;
                delete parsedObject.Fields;
                delete parsedObject.EnableContentTypes;
                delete parsedObject.Views;
                delete parsedObject.Items;
                delete parsedObject.TemplateType;
                delete parsedObject.Files;
                delete parsedObject.ContentTypeBindings;
                break;
            default:
                delete parsedObject.InternalName;
                delete parsedObject.EnableContentTypes;
                delete parsedObject.ControlOption;
                delete parsedObject.Fields;
                delete parsedObject.Views;
                delete parsedObject.Title;
                delete parsedObject.Items;
                delete parsedObject.Description;
                delete parsedObject.TemplateType;
                delete parsedObject.Files;
                delete parsedObject.ContentTypeBindings;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}
