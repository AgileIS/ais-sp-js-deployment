import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Folders, Folder } from "@agileis/sp-pnp-js/lib/sharepoint/rest/folders";
import { Files, File, NodeFile } from "@agileis/sp-pnp-js/lib/sharepoint/rest/files";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { Item } from "@agileis/sp-pnp-js/lib/sharepoint/rest/items";
import { ISPObjectHandler } from "../interfaces/ObjectHandler/ispobjecthandler";
import { IFile } from "../interfaces/Types/IFile";
import { IFolder } from "../interfaces/Types/IFolder";
import { IItem } from "../interfaces/Types/IItem";
import { ControlOption } from "../Constants/ControlOption";
import { IPromiseResult } from "../interfaces/IPromiseResult";
import { Util } from "../Util/Util";
import * as fs from "fs";
import * as mime from "mime";

export class FileHandler implements ISPObjectHandler {
    public execute(fileFolderConfig: IFile | IFolder, parentPromise: Promise<IPromiseResult<Web | Folder | List>>): Promise<IPromiseResult<File | Folder | void>> {
        return new Promise<IPromiseResult<File | Folder | void>>((resolve, reject) => {
            parentPromise.then(parentResult => {
                parentResult.value.get()
                    .then(parentRequestResult => {
                        let parent = parentResult.value;
                        if (parentResult.value instanceof Web) {
                            parent = new Folder(parentResult.value.toUrl(), (fileFolderConfig as IFile).Src ? "" : `GetFolderByServerRelativeUrl('${fileFolderConfig.Name}')`);
                             Util.Resolve<Folder>(resolve, fileFolderConfig.Name, `'${fileFolderConfig.Name}' is RootFolder`, parent as Folder);
                        } else {
                            if (parentResult.value instanceof List) {
                                parent = new Folder(parentRequestResult.RootFolder.__deferred.uri);
                            }

                            let processing: Promise<IPromiseResult<File | Folder | void>>;
                            if ((fileFolderConfig as IFile).Src) {
                                processing = this.processingFileConfig(fileFolderConfig as IFile, (parent as Folder).files);
                            } else {
                                processing = this.processingFolderConfig(fileFolderConfig as IFolder, (parent as Folder).folders);
                            }
                            processing
                                .then((fileFolderProsssingResult) => { resolve(fileFolderProsssingResult); })
                                .catch((error) => { reject(error); });
                        }

                    })
                    .catch(error => {
                        Util.Reject<void>(reject, fileFolderConfig.Name, `Error while requesting parent ('${parentResult.value.toUrl()}') for element: '${fileFolderConfig.Name}': ` + error);
                    });
            });
        });
    }

    private processingFolderConfig(folderConfig: IFolder, parentFolder: Folders): Promise<IPromiseResult<Folder | void>> {
        return new Promise<IPromiseResult<Folder | void>>((resolve, reject) => {
            let processingText = folderConfig.ControlOption === ControlOption.Add || folderConfig.ControlOption === undefined || folderConfig.ControlOption === ""
                ? "Add" : folderConfig.ControlOption;
            Logger.write(`Processing ${processingText} folder: '${folderConfig.Name}' to ${parentFolder.toUrl()}`, Logger.LogLevel.Info);

            let folder = parentFolder.getByName(folderConfig.Name);
            folder.get()
                .then(folderRequestResult => {
                    switch (folderConfig.ControlOption) {
                        case ControlOption.Delete:
                            this.deleteFolder(folderConfig, folder)
                                .then((folderProcessingResult) => { resolve(folderProcessingResult); })
                                .catch((error) => { reject(error); });
                            break;
                        case ControlOption.Update:
                        default: // tslint:disable-line
                            Util.Resolve<Folder>(resolve, folderConfig.Name, `Folder with the name '${folderConfig.Name}' already exists in '${parentFolder.toUrl()}'`, folder);
                            break;
                    }
                })
                .catch((error) => {
                    if (error === "Error making GET request: Not Found") {
                        switch (folderConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Reject<void>(reject, folderConfig.Name, `Folder with Name '${folderConfig.Name}' does not exists in '${folder.parentFolder}'`);
                                break;
                            case ControlOption.Update:
                            default: // tslint:disable-line
                                this.addFolder(folderConfig, parentFolder)
                                    .then((folderProcessingResult) => { resolve(folderProcessingResult); })
                                    .catch((addingError) => { reject(addingError); });
                                break;
                        }
                    } else {
                        Util.Reject<void>(reject, folderConfig.Name, `Error while requesting folder with the title '${folderConfig.Name}' from parent '${folder.parentFolder}': ` + error);
                    }
                });
        });
    }

    private processingFileConfig(fileConfig: IFile, parentFolder: Files): Promise<IPromiseResult<File>> {
        return new Promise<IPromiseResult<File>>((resolve, reject) => {
            let processingText = fileConfig.ControlOption === ControlOption.Add || fileConfig.ControlOption === undefined || fileConfig.ControlOption === ""
                ? "Add" : fileConfig.ControlOption;
            Logger.write(`Processing ${processingText} file: '${fileConfig.Name}' in ${parentFolder.toUrl()}`, Logger.LogLevel.Info);

            let file = parentFolder.getByName(fileConfig.Name);
            file.get()
                .then(folderRequestResult => {
                    let processingPromise: Promise<IPromiseResult<File>> = undefined;
                    switch (fileConfig.ControlOption) {
                        case ControlOption.Delete:
                            processingPromise = this.deleteFile(fileConfig, file);
                            break;
                        case ControlOption.Update:

                            break;
                        default:
                            Util.Resolve<File>(resolve, fileConfig.Name, `File with the name '${fileConfig.Name}' already exists in '${parentFolder.toUrl()}'`, file);
                            break;
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((fileProcessingResult) => { resolve(fileProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else {
                        Logger.write("File handler processing promise is undefined!");
                    }
                })
                .catch((error) => {
                    if (error === "Error making GET request: Not Found") {
                        switch (fileConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Reject<void>(reject, fileConfig.Name, `File with Name '${fileConfig.Name}' does not exists in '${parentFolder.toUrl()}'`);
                                break;
                            case ControlOption.Update:
                            default: // tslint:disable-line
                                this.addFile(fileConfig, parentFolder)
                                    .then((folderProcessingResult) => {
                                        resolve(folderProcessingResult);
                                        // ToDo update FileItem properties
                                    })
                                    .catch((addingError) => { reject(addingError); });
                                break;
                        }
                    } else {
                        Util.Reject<void>(reject, fileConfig.Name, `Error while requesting file with the title '${fileConfig.Name}' from parent '${parentFolder.toUrl()}': ` + error);
                    }
                });
        });
    }

    private addFolder(folderConfig: IFolder, parentFolder: Folders): Promise<IPromiseResult<Folder>> {
        return new Promise<IPromiseResult<Folder>>((resolve, reject) => {
            parentFolder.add(folderConfig.Name)
                .then((folderAddResult) => { Util.Resolve<Folder>(resolve, folderConfig.Name, `Added item: '${folderConfig.Name}'`, folderAddResult.folder); })
                .catch((error) => { Util.Reject<void>(reject, folderConfig.Name, `Error while adding folder with name '${folderConfig.Name}': ` + error); });
        });
    }

    private deleteFolder(folderConfig: IFolder, folder: Folder): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            folder.delete()
                .then(() => { Util.Resolve<void>(resolve, folderConfig.Name, `Deleted folder: '${folderConfig.Name}'`); })
                .catch((error) => { Util.Reject<void>(reject, `Error while deleting folder with name '${folderConfig.Name}': ` + error, folderConfig.Name); });
        });
    }

    private addFile(fileConfig: IFile, parentFolder: Files): Promise<IPromiseResult<File>> {
        return new Promise<IPromiseResult<File>>((resolve, reject) => {
            let file: NodeFile = {
                data: fs.readFileSync(fileConfig.Src),
                mime: mime.lookup(fileConfig.Name),
            };
            parentFolder.add(fileConfig.Name, file)
                .then((fileAddResult) => { Util.Resolve<File>(resolve, fileConfig.Name, `Added file: '${fileConfig.Name}'`, fileAddResult.file); })
                .catch((error) => { Util.Reject<void>(reject, fileConfig.Name, `Error while adding folder with name '${fileConfig.Name}': ` + error); });
        });
    }

    private deleteFile(fileConfig: IFile, file: File): Promise<IPromiseResult<File>> {
        return new Promise<IPromiseResult<File>>((resolve, reject) => {
            file.delete()
                .then(() => { Util.Resolve(resolve, fileConfig.Name, `Deleted file: '${fileConfig.Name}'`); })
                .catch((error) => { Util.Reject(reject, fileConfig.Name, `Error while deleting file with name '${fileConfig.Name}': ` + error); });
        });
    }

    private updateFile(fileConfig: IFile, file: File): Promise<IPromiseResult<Item>> {
        return new Promise<IPromiseResult<Item>>((resolve, reject) => {
            let properties = this.createProperties(fileConfig.Properties as IItem);
            file.listItemAllFields.update(properties)
                .then((itemUpdateResult) => { Util.Resolve<Item>(resolve, fileConfig.Name, `Updated item: '${fileConfig.Name}'`, itemUpdateResult.item); })
                .catch((error) => { Util.Reject(reject, `Error while updating item with title '${fileConfig.Name}': ` + error, fileConfig.Name); });
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
                delete parsedObject.Title;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}