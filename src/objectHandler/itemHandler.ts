import { Logger } from "ais-sp-pnp-js/lib/utils/logging";
import { List } from "ais-sp-pnp-js/lib/sharepoint/rest/lists";
import { Item } from "ais-sp-pnp-js/lib/sharepoint/rest/items";
import { ISPObjectHandler } from "../interfaces/objectHandler/iSpObjectHandler";
import { IItem } from "../interfaces/types/iItem";
import { IPromiseResult } from "../interfaces/iPromiseResult";
import { ControlOption } from "../constants/controlOption";
import { Util } from "../util/util";

export class ItemHandler implements ISPObjectHandler {
    private handlerName = "ItemHandler";
    public execute(itemConfig: IItem, parentPromise: Promise<IPromiseResult<List>>): Promise<IPromiseResult<void | Item>> {
        return new Promise<IPromiseResult<void | Item>>((resolve, reject) => {
            parentPromise.then(promiseResult => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, this.handlerName,
                        `Item handler parent promise value result is null or undefined for the item with the title '${itemConfig.Title}'!`);
                } else {
                    let list = promiseResult.value;
                    Util.tryToProcess(itemConfig.Title, () => { return this.processingItemConfig(itemConfig, list); }, this.handlerName)
                        .then((itemProcessingResult) => { resolve(itemProcessingResult); })
                        .catch((error) => { reject(error); });
                }
            });
        });
    }

    private processingItemConfig(itemConfig: IItem, list: List): Promise<IPromiseResult<void | Item>> {
        return new Promise<IPromiseResult<void | Item>>((resolve, reject) => {
            let processingText = itemConfig.ControlOption === ControlOption.ADD || itemConfig.ControlOption === undefined || itemConfig.ControlOption === ""
                ? "Add" : itemConfig.ControlOption;
            Logger.write(`${this.handlerName} - Processing ${processingText} item: '${itemConfig.Title}'.`, Logger.LogLevel.Info);

            list.items.filter(`Title eq '${itemConfig.Title}'`).select("Id").get()
                .then((itemRequestResults) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | Item>> = undefined;

                    if (itemRequestResults && itemRequestResults.length === 1) {
                        Logger.write(`${this.handlerName} - Found Item with the title: '${itemConfig.Title}'`, Logger.LogLevel.Info);
                        let item = list.items.getById(itemRequestResults[0].Id);
                        switch (itemConfig.ControlOption) {
                            case ControlOption.UPDATE:
                                processingPromise = this.updateItem(itemConfig, item);
                                break;
                            case ControlOption.DELETE:
                                processingPromise = this.deleteItem(itemConfig, item);
                                break;
                            default:
                                Util.Resolve<Item>(resolve, this.handlerName, `Item with the title '${itemConfig.Title}' does not have to be added, because it already exists.`, item);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (itemConfig.ControlOption) {
                            case ControlOption.DELETE:
                                Util.Resolve<void>(resolve, this.handlerName, `Item with Title '${itemConfig.Title}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.UPDATE:
                                itemConfig.ControlOption = ControlOption.ADD;
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
                        Logger.write(`${this.handlerName} - Processing promise is undefined!`, Logger.LogLevel.Error);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while requesting item with the title '${itemConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private addItem(itemConfig: IItem, parentList: List): Promise<IPromiseResult<Item>> {
        return new Promise<IPromiseResult<Item>>((resolve, reject) => {
            let properties = this.createProperties(itemConfig);
            parentList.items.add(properties)
                .then((itemAddResult) => { Util.Resolve<Item>(resolve, this.handlerName, `Added item: '${itemConfig.Title}'.`, itemAddResult.item); })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while adding item with title '${itemConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private updateItem(itemConfig: IItem, item: Item): Promise<IPromiseResult<Item>> {
        return new Promise<IPromiseResult<Item>>((resolve, reject) => {
            let properties = this.createProperties(itemConfig);
            item.update(properties)
                .then((itemUpdateResult) => { Util.Resolve<Item>(resolve, this.handlerName, `Updated item: '${itemConfig.Title}'.`, itemUpdateResult.item); })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while updating item with title '${itemConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private deleteItem(itemConfig: IItem, item: Item): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            item.delete()
                .then(() => { Util.Resolve<void>(resolve, this.handlerName, `Deleted item: '${itemConfig.Title}'.`); })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while deleting item with title '${itemConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private createProperties(itemConfig: IItem) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(itemConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (itemConfig.ControlOption) {
            case ControlOption.UPDATE:
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
