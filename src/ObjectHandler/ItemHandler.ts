import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { Item } from "@agileis/sp-pnp-js/lib/sharepoint/rest/Items";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IItem } from "../Interfaces/Types/IItem";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { ControlOption } from "../Constants/ControlOption";
import { Util } from "../Util/Util";

export class ItemHandler implements ISPObjectHandler {
    public execute(itemConfig: IItem, parentPromise: Promise<IPromiseResult<List>>): Promise<IPromiseResult<void | Item>> {
        return new Promise<IPromiseResult<void | Item>>((resolve, reject) => {
            parentPromise.then(promiseResult => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, itemConfig.Title,
                        `Item handler parent promise value result is null or undefined for the item with the title '${itemConfig.Title}'!`);
                } else {
                    let list = promiseResult.value;
                    this.processingItemConfig(itemConfig, list)
                        .then((itemProsssingResult) => { resolve(itemProsssingResult); })
                        .catch((error) => {
                            Util.Retry(error, itemConfig.Title,
                                () => {
                                    return this.processingItemConfig(itemConfig, list);
                                });
                        });
                }
            });
        });
    }

    private processingItemConfig(itemConfig: IItem, list: List): Promise<IPromiseResult<void | Item>> {
        return new Promise<IPromiseResult<void | Item>>((resolve, reject) => {
            let processingText = itemConfig.ControlOption === ControlOption.Add || itemConfig.ControlOption === undefined || itemConfig.ControlOption === ""
                ? "Add" : itemConfig.ControlOption;
            Logger.write(`Processing ${processingText} item: '${itemConfig.Title}'.`, Logger.LogLevel.Info);

            list.items.filter(`Title eq '${itemConfig.Title}'`).select("Id").get()
                .then((itemRequestResults) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | Item>> = undefined;

                    if (itemRequestResults && itemRequestResults.length === 1) {
                        Logger.write(`Found Item with the title: '${itemConfig.Title}'`);
                        let item = list.items.getById(itemRequestResults[0].Id);
                        switch (itemConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateItem(itemConfig, item);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteItem(itemConfig, item);
                                break;
                            default:
                                Util.Resolve<Item>(resolve, itemConfig.Title, `Item with the title '${itemConfig.Title}' does not have to be added, because it already exists.`, item);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (itemConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Resolve<void>(resolve, itemConfig.Title, `Item with Title '${itemConfig.Title}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.Update:
                                itemConfig.ControlOption = ControlOption.Add;
                            default:
                                processingPromise = this.addItem(itemConfig, list);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((itemProcessingResult) => { resolve(itemProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write("List handler processing promise is undefined!", Logger.LogLevel.Error);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, itemConfig.Title, `Error while requesting item with the title '${itemConfig.Title}': ` + error); });
        });
    }

    private addItem(itemConfig: IItem, parentList: List): Promise<IPromiseResult<Item>> {
        return new Promise<IPromiseResult<Item>>((resolve, reject) => {
            let properties = this.createProperties(itemConfig);
            parentList.items.add(properties)
                .then((itemAddResult) => { Util.Resolve<Item>(resolve, itemConfig.Title, `Added item: '${itemConfig.Title}'.`, itemAddResult.item); })
                .catch((error) => { Util.Reject<void>(reject, itemConfig.Title, `Error while adding item with title '${itemConfig.Title}': ` + error); });
        });
    }

    private updateItem(itemConfig: IItem, item: Item): Promise<IPromiseResult<Item>> {
        return new Promise<IPromiseResult<Item>>((resolve, reject) => {
            let properties = this.createProperties(itemConfig);
            item.update(properties)
                .then((itemUpdateResult) => { Util.Resolve<Item>(resolve, itemConfig.Title, `Updated item: '${itemConfig.Title}'.`, itemUpdateResult.item); })
                .catch((error) => { Util.Reject<void>(reject, itemConfig.Title, `Error while updating item with title '${itemConfig.Title}': ` + error); });
        });
    }

    private deleteItem(itemConfig: IItem, item: Item): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            item.delete()
                .then(() => { Util.Resolve<void>(resolve, itemConfig.Title, `Deleted item: '${itemConfig.Title}'.`); })
                .catch((error) => { Util.Reject<void>(reject, itemConfig.Title, `Error while deleting item with title '${itemConfig.Title}': ` + error); });
        });
    }

    private createProperties(itemConfig: IItem) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(itemConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (itemConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.ControlOption;
                break;
            default:
                delete parsedObject.ControlOption;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}
