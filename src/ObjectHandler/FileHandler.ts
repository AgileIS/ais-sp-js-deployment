import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Folders, Folder } from "@agileis/sp-pnp-js/lib/sharepoint/rest/folders";
import { Files, File} from "@agileis/sp-pnp-js/lib/sharepoint/rest/files";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IFile } from "../interface/Types/IFile";
import { IFolder } from "../interface/Types/IFolder";
import { IItem } from "../interface/Types/IItem";
import { ControlOption } from "../Constants/ControlOption";
import { Resolve, Reject } from "../Util/Util";
import * as fs from "fs";
import * as mime from "mime";

export class FileHandler implements ISPObjectHandler {
    public execute(fileFolderConfig: IFile | IFolder, parentPromise: Promise<Web | Folder | List>): Promise<File | Folder> {
        return new Promise<File | Folder>((resolve, reject) => {
            parentPromise.then(parentResult => {
                parentResult.get()
                    .then(parentRequestResult => {
                        let parent = parentResult;
                        if (parentResult instanceof Web) {
                            parent = new Folder(parentResult.toUrl(), (fileFolderConfig as IFile).Src ? "" : `GetFolderByServerRelativeUrl('${fileFolderConfig.Name}')`);
                            resolve(parent as Folder);
                        } else {
                            if (parentResult instanceof List) {
                                parent = new Folder(parentRequestResult.RootFolder.__deferred.uri);
                            }

                            let processing: Promise<File | Folder>;
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
                        Reject(reject, `Error while requesting parent ('${parentResult.toUrl()}') for element with the title '${fileFolderConfig.Name}': ` + error, fileFolderConfig.Name);
                    });
            });
        });
    }

    private processingFolderConfig(folderConfig: IFolder, parentFolder: Folders): Promise<Folder> {
        return new Promise<File | Folder>((resolve, reject) => {
            let processingText = folderConfig.ControlOption === ControlOption.Add || folderConfig.ControlOption === undefined || folderConfig.ControlOption === ""
                ? "Add" : folderConfig.ControlOption;
            Logger.write(`Processing ${processingText} folder: '${folderConfig.Name}' to ${parentFolder.toUrl()}`, Logger.LogLevel.Info);

            let folder = parentFolder.getByName(folderConfig.Name);
            folder.get()
                .then(folderRequestResult => {
                    let processingPromise: Promise<Folder> = undefined;
                    switch (folderConfig.ControlOption) {
                        case ControlOption.Delete:
                            processingPromise = this.deleteFolder(folderConfig, folder);
                            break;
                        case ControlOption.Update:
                        default: // tslint:disable-line
                            Resolve(resolve, `Folder with the name '${folderConfig.Name}' already exists in '${parentFolder.toUrl()}'`, folderConfig.Name, folder);
                            break;
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((folderProcessingResult) => { resolve(folderProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else {
                        Logger.write("File handler processing promise is undefined!");
                    }
                })
                .catch((error) => {
                    if (error === "Error making GET request: Not Found") {
                        switch (folderConfig.ControlOption) {
                            case ControlOption.Delete:
                                Reject(reject, `Folder with Name '${folderConfig.Name}' does not exists in '${folder.parentFolder}'`, folderConfig.Name);
                                break;
                            case ControlOption.Update:
                            default: // tslint:disable-line
                                this.addFolder(folderConfig, parentFolder)
                                    .then((folderProcessingResult) => { resolve(folderProcessingResult); })
                                    .catch((addingError) => { reject(addingError); });
                                break;
                        }
                    } else {
                        Reject(reject, `Error while requesting folder with the title '${folderConfig.Name}' from parent '${folder.parentFolder}': ` + error, folderConfig.Name);
                    }
                });
        });
    }

    private processingFileConfig(fileConfig: IFile, parentFolder: Files): Promise<File> {
        return new Promise<File | Folder>((resolve, reject) => {
            let processingText = fileConfig.ControlOption === ControlOption.Add || fileConfig.ControlOption === undefined || fileConfig.ControlOption === ""
                ? "Add" : fileConfig.ControlOption;
            Logger.write(`Processing ${processingText} file: '${fileConfig.Name}' in ${parentFolder.toUrl()}`, Logger.LogLevel.Info);

            let file = parentFolder.getByName(fileConfig.Name);
            file.get()
                .then(folderRequestResult => {
                    let processingPromise: Promise<File> = undefined;
                    switch (fileConfig.ControlOption) {
                        case ControlOption.Delete:
                            processingPromise = this.deleteFile(fileConfig, file);
                            break;
                        case ControlOption.Update:

                            break;
                        default:
                            Resolve(resolve, `File with the name '${fileConfig.Name}' already exists in '${parentFolder.toUrl()}'`, fileConfig.Name, file);
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
                                Reject(reject, `File with Name '${fileConfig.Name}' does not exists in '${parentFolder.toUrl()}'`, fileConfig.Name);
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
                        Reject(reject, `Error while requesting file with the title '${fileConfig.Name}' from parent '${parentFolder.toUrl()}': ` + error, fileConfig.Name);
                    }
                });
        });
    }

    private addFolder(folderConfig: IFolder, parentFolder: Folders): Promise<Folder> {
        return new Promise<Folder>((resolve, reject) => {
            parentFolder.add(folderConfig.Name)
                .then((folderAddResult) => { Resolve(resolve, `Added item: '${folderConfig.Name}'`, folderConfig.Name, folderAddResult.folder); })
                .catch((error) => { Reject(reject, `Error while adding folder with name '${folderConfig.Name}': ` + error, folderConfig.Name); });
        });
    }

    private deleteFolder(folderConfig: IFolder, folder: Folder): Promise<Folder> {
        return new Promise<Folder>((resolve, reject) => {
            folder.delete()
                .then(() => { Resolve(resolve, `Deleted folder: '${folderConfig.Name}'`, folderConfig.Name); })
                .catch((error) => { Reject(reject, `Error while deleting folder with name '${folderConfig.Name}': ` + error, folderConfig.Name); });
        });
    }

    private addFile(fileConfig: IFile, parentFolder: Files): Promise<File> {
        return new Promise<File>((resolve, reject) => {
            let file = fs.readFileSync(fileConfig.Src);
            parentFolder.add(fileConfig.Name, file, mime.lookup(fileConfig.Name))
                .then((fileAddResult) => { Resolve(resolve, `Added file: '${fileConfig.Name}'`, fileConfig.Name, fileAddResult.file); })
                .catch((error) => { Reject(reject, `Error while adding folder with name '${fileConfig.Name}': ` + error, fileConfig.Name); });
        });
    }

    private deleteFile(fileConfig: IFile, file: File): Promise<File> {
        return new Promise<File>((resolve, reject) => {
            file.delete()
                .then(() => { Resolve(resolve, `Deleted file: '${fileConfig.Name}'`, fileConfig.Name); })
                .catch((error) => { Reject(reject, `Error while deleting file with name '${fileConfig.Name}': ` + error, fileConfig.Name); });
        });
    }

    private updateFile(fileConfig: IFile, file: File): Promise<File> {
        return new Promise<File>((resolve, reject) => {
            let properties = this.createProperties(fileConfig.Properties as IItem);
            file.listItemAllFields.update(properties)
                .then((itemUpdateResult) => { Resolve(resolve, `Updated item: '${fileConfig.Name}'`, fileConfig.Name, itemUpdateResult.item); })
                .catch((error) => { Reject(reject, `Error while updating item with title '${fileConfig.Name}': ` + error, fileConfig.Name); });
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
