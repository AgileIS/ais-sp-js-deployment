import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { Item } from "@agileis/sp-pnp-js/lib/sharepoint/rest/Items";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IItem } from "../interface/Types/IItem";
import { ControlOption } from "../Constants/ControlOption";
import { Resolve, Reject, Retry } from "../Util/Util";

export class ItemHandler implements ISPObjectHandler {
    public execute(itemConfig: IItem, parentPromise: Promise<List>): Promise<Item> {
        return new Promise<Item>((resolve, reject) => {
            parentPromise.then(parentList => {
                this.processingItemConfig(itemConfig, parentList)
                    .then((itemProsssingResult) => { resolve(itemProsssingResult); })
                    // .catch((error) => { reject(error); });
                    .catch(() => { Retry(() => {
                        return this.processingItemConfig(itemConfig, parentList);
                    }, itemConfig.Title); });
            });
        });
    }

    private processingItemConfig(itemConfig: IItem, parentList: List): Promise<Item> {
        return new Promise<Item>((resolve, reject) => {
            let processingText = itemConfig.ControlOption === ControlOption.Add || itemConfig.ControlOption === undefined || itemConfig.ControlOption === ""
                ? "Add" : itemConfig.ControlOption;
            Logger.write(`Processing ${processingText} item: '${itemConfig.Title}'`, Logger.LogLevel.Info);

            parentList.items.filter(`Title eq '${itemConfig.Title}'`).select("Id").get()
                .then((itemRequestResults) => {
                    let processingPromise: Promise<Item> = undefined;

                    if (itemRequestResults && itemRequestResults.length === 1) {
                        let item = parentList.items.getById(itemRequestResults[0].Id);
                        switch (itemConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateItem(itemConfig, item);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteItem(itemConfig, item);
                                break;
                            default:
                                Resolve(resolve, `Item with the title '${itemConfig.Title}' already exists`, itemConfig.Title, item);
                                break;
                        }
                    } else {
                        switch (itemConfig.ControlOption) {
                            case ControlOption.Delete:
                                Resolve(resolve, `Item with the title '${itemConfig.Title}'  not found anyway`, itemConfig.Title);
                                break;
                            case ControlOption.Update:
                            default: // tslint:disable-line
                                processingPromise = this.addItem(itemConfig, parentList);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((itemProcessingResult) => { resolve(itemProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else {
                        Logger.write("List handler processing promise is undefined!");
                    }
                })
                .catch((error) => { Reject(reject, `Error while requesting item with the title '${itemConfig.Title}': ` + error, itemConfig.Title); });
        });
    }

    private addItem(itemConfig: IItem, parentList: List): Promise<Item> {
        return new Promise<Item>((resolve, reject) => {
            let properties = this.createProperties(itemConfig);
            parentList.items.add(properties)
                .then((itemAddResult) => { Resolve(resolve, `Added item: '${itemConfig.Title}'`, itemConfig.Title, itemAddResult.item); })
                .catch((error) => { Reject(reject, `Error while adding item with title '${itemConfig.Title}': ` + error, itemConfig.Title); });
        });
    }

    private updateItem(itemConfig: IItem, item: Item): Promise<Item> {
        return new Promise<Item>((resolve, reject) => {
            let properties = this.createProperties(itemConfig);
            item.update(properties)
                .then((itemUpdateResult) => { Resolve(resolve, `Updated item: '${itemConfig.Title}'`, itemConfig.Title, itemUpdateResult.item); })
                .catch((error) => { Reject(reject, `Error while updating item with title '${itemConfig.Title}': ` + error, itemConfig.Title); });
        });
    }

    private deleteItem(itemConfig: IItem, item: Item): Promise<Item> {
        return new Promise<Item>((resolve, reject) => {
            item.delete()
                .then(() => { Resolve(resolve, `Deleted item: '${itemConfig.Title}'`, itemConfig.Title); })
                .catch((error) => { Reject(reject, `Error while deleting item with title '${itemConfig.Title}': ` + error, itemConfig.Title); });
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
