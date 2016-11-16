import { Logger } from "ais-sp-pnp-js/lib/utils/logging";
import { Web } from "ais-sp-pnp-js/lib/sharepoint/rest/webs";
import { Folders, Folder } from "ais-sp-pnp-js/lib/sharepoint/rest/folders";
import { Files, File, NodeFile } from "ais-sp-pnp-js/lib/sharepoint/rest/files";
import { List } from "ais-sp-pnp-js/lib/sharepoint/rest/lists";
import { Item } from "ais-sp-pnp-js/lib/sharepoint/rest/items";
import { ISPObjectHandler } from "../interfaces/objectHandler/iSpObjectHandler";
import { IFile } from "../interfaces/types/iFile";
import { IFolder } from "../interfaces/types/iFolder";
import { IItem } from "../interfaces/types/iItem";
import { ControlOption } from "../constants/controlOption";
import { IPromiseResult } from "../interfaces/iPromiseResult";
import { Util } from "../util/util";
import { spawn } from "child_process";
import * as fs from "fs";
import * as mime from "mime";

declare var window: Window;
interface Window { // tslint:disable-line
    _spPageContextInfo: any;
}

export class FileHandler implements ISPObjectHandler {
    private handlerName = "FileHandler";
    public execute(fileFolderConfig: IFile | IFolder, parentPromise: Promise<IPromiseResult<Web | Folder | List>>): Promise<IPromiseResult<File | Folder | void>> {
        return new Promise<IPromiseResult<File | Folder | void>>((resolve, reject) => {
            parentPromise.then(parentResult => {
                parentResult.value.get()
                    .then(parentRequestResult => {
                        let parent = parentResult.value;
                        if (parentResult.value instanceof Web) {
                            parent = new Folder(parentResult.value.toUrl(), (fileFolderConfig as IFile).Src ? "" : `GetFolderByServerRelativeUrl('${fileFolderConfig.Name}')`);
                            Util.Resolve<Folder>(resolve, this.handlerName, `'${fileFolderConfig.Name}' is RootFolder`, parent as Folder);
                        } else {
                            if (parentResult.value instanceof List) {
                                parent = new Folder(parentRequestResult.RootFolder.__deferred.uri);
                            }
                            Util.tryToProcess(fileFolderConfig.Name, () => {
                                let processing: Promise<IPromiseResult<File | Folder | void>>;
                                if ((fileFolderConfig as IFile).Src) {
                                    processing = this.processingFileConfig(fileFolderConfig as IFile, (parent as Folder).files);
                                } else {
                                    processing = this.processingFolderConfig(fileFolderConfig as IFolder, (parent as Folder).folders);
                                }
                                return processing;

                            }, this.handlerName)
                                .then((fileFolderProcessingResult) => { resolve(fileFolderProcessingResult); })
                                .catch((error) => { reject(error); });
                        }

                    })
                    .catch(error => {
                        Util.Reject<void>(reject, this.handlerName,
                            `Error while requesting parent ('${parentResult.value.toUrl()}') for element: '${fileFolderConfig.Name}': ` + Util.getErrorMessage(error));
                    });
            });
        });
    }

    private processingFolderConfig(folderConfig: IFolder, parentFolder: Folders): Promise<IPromiseResult<Folder | void>> {
        return new Promise<IPromiseResult<Folder | void>>((resolve, reject) => {
            let processingText = folderConfig.ControlOption === ControlOption.ADD || folderConfig.ControlOption === undefined || folderConfig.ControlOption === ""
                ? "Add" : folderConfig.ControlOption;
            Logger.write(`${this.handlerName} - Processing ${processingText} folder: '${folderConfig.Name}' to ${parentFolder.toUrl()}`, Logger.LogLevel.Info);

            let folder = parentFolder.getByName(folderConfig.Name);
            folder.get()
                .then(folderRequestResult => {
                    switch (folderConfig.ControlOption) {
                        case ControlOption.DELETE:
                            this.deleteFolder(folderConfig, folder)
                                .then((folderProcessingResult) => { resolve(folderProcessingResult); })
                                .catch((error) => { reject(error); });
                            break;
                        case ControlOption.UPDATE:
                        default: // tslint:disable-line
                            Util.Resolve<Folder>(resolve, this.handlerName, `Folder with the name '${folderConfig.Name}' already exists in '${parentFolder.toUrl()}'`, folder);
                            break;
                    }
                })
                .catch((error) => {
                    if (error === "Error making GET request: Not Found") {
                        switch (folderConfig.ControlOption) {
                            case ControlOption.DELETE:
                                Util.Reject<void>(reject, this.handlerName, `Folder with Name '${folderConfig.Name}' does not exists in '${folder.parentFolder}'`);
                                break;
                            case ControlOption.UPDATE:
                            default: // tslint:disable-line
                                this.addFolder(folderConfig, parentFolder)
                                    .then((folderProcessingResult) => { resolve(folderProcessingResult); })
                                    .catch((addingError) => { reject(addingError); });
                                break;
                        }
                    } else {
                        Util.Reject<void>(reject, this.handlerName,
                            `Error while requesting folder with the title '${folderConfig.Name}' from parent '${folder.parentFolder}': ` + Util.getErrorMessage(error));
                    }
                });
        });
    }

    private processingFileConfig(fileConfig: IFile, parentFolder: Files): Promise<IPromiseResult<File | void>> {
        return new Promise<IPromiseResult<File | void>>((resolve, reject) => {
            let processingText = fileConfig.ControlOption === ControlOption.ADD || fileConfig.ControlOption === undefined || fileConfig.ControlOption === ""
                ? "Add" : fileConfig.ControlOption;
            Logger.write(`${this.handlerName} - Processing ${processingText} file: '${fileConfig.Name}' in ${parentFolder.toUrl()}`, Logger.LogLevel.Info);

            let file = parentFolder.getByName(fileConfig.Name);
            file.get()
                .then(folderRequestResult => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<File | void>> = undefined;
                    switch (fileConfig.ControlOption) {
                        case ControlOption.DELETE:
                            processingPromise = this.deleteFile(fileConfig, file);
                            break;
                        case ControlOption.UPDATE:
                            processingPromise = this.addFile(fileConfig, parentFolder, true, this.resolvePreTask(fileConfig));
                            break;
                        default:
                            Util.Resolve<File>(resolve, this.handlerName, `File with the name '${fileConfig.Name}' already exists in '${parentFolder.toUrl()}'`, file);
                            rejectOrResolved = true;
                            break;
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((fileProcessingResult) => { resolve(fileProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write(`${this.handlerName} - Processing promise is undefined!`, Logger.LogLevel.Error);
                    }
                })
                .catch((error) => {
                    if (error === "Error making GET request: Not Found") {
                        switch (fileConfig.ControlOption) {
                            case ControlOption.DELETE:
                                Util.Reject<void>(reject, this.handlerName, `File with Name '${fileConfig.Name}' does not exists in '${parentFolder.toUrl()}'`);
                                break;
                            case ControlOption.UPDATE:
                            default: // tslint:disable-line
                                this.addFile(fileConfig, parentFolder, fileConfig.Overwrite === true, this.resolvePreTask(fileConfig))
                                    .then((folderProcessingResult) => {
                                        resolve(folderProcessingResult);
                                    })
                                    .catch((addingError) => { reject(addingError); });
                                break;
                        }
                    } else {
                        Util.Reject<void>(reject, this.handlerName,
                            `Error while requesting file with the title '${fileConfig.Name}' from parent '${parentFolder.toUrl()}': ` + Util.getErrorMessage(error));
                    }
                });
        });
    }

    private addFolder(folderConfig: IFolder, parentFolder: Folders): Promise<IPromiseResult<Folder>> {
        return new Promise<IPromiseResult<Folder>>((resolve, reject) => {
            parentFolder.add(folderConfig.Name)
                .then((folderAddResult) => { Util.Resolve<Folder>(resolve, this.handlerName, `Added item: '${folderConfig.Name}'`, folderAddResult.folder); })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while adding folder with name '${folderConfig.Name}': ` + Util.getErrorMessage(error)); });
        });
    }

    private deleteFolder(folderConfig: IFolder, folder: Folder): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            folder.delete()
                .then(() => { Util.Resolve<void>(resolve, this.handlerName, `Deleted folder: '${folderConfig.Name}'`); })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while deleting folder with name '${folderConfig.Name}': ` + Util.getErrorMessage(error)); });
        });
    }

    private addFile(fileConfig: IFile, parentFolder: Files, overwrite: boolean, preTask: Promise<any> = Promise.resolve()): Promise<IPromiseResult<File>> {
        return new Promise<IPromiseResult<File>>((resolve, reject) => {
            preTask
                .then(() => {
                    let file: NodeFile = {
                        data: fs.readFileSync(fileConfig.Src),
                        mime: mime.lookup(fileConfig.Name),
                    };
                    parentFolder.add(fileConfig.Name, file, overwrite)
                        .then((fileAddResult) => {
                            if (fileConfig.Properties) {
                                this.updateFileProperties(fileConfig, fileAddResult.file)
                                    .then((fileUpdateResult) => {
                                        Util.Resolve<File>(resolve, this.handlerName, `Added file: '${fileConfig.Name}'`, fileAddResult.file);
                                    })
                                    .catch((error) => {
                                        Util.Reject<void>(reject, this.handlerName, `Error while updating file item fields with name '${fileConfig.Name}': ` + Util.getErrorMessage(error));
                                    });
                            } else {
                                Util.Resolve<File>(resolve, this.handlerName, `Added file: '${fileConfig.Name}'`, fileAddResult.file);
                            }

                        })
                        .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while adding file with name '${fileConfig.Name}': ` + Util.getErrorMessage(error)); });
                })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while proccesing preTask for file with name '${fileConfig.Name}': ` + Util.getErrorMessage(error)); });
        });
    }

    private deleteFile(fileConfig: IFile, file: File): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            file.delete()
                .then(() => { Util.Resolve<void>(resolve, this.handlerName, `Deleted file: '${fileConfig.Name}'`); })
                .catch((error) => { Util.Reject(reject, this.handlerName, `Error while deleting file with name '${fileConfig.Name}': ` + Util.getErrorMessage(error)); });
        });
    }

    private updateFileProperties(fileConfig: IFile, file: File): Promise<IPromiseResult<Item>> {
        return new Promise<IPromiseResult<Item>>((resolve, reject) => {
            let properties = this.createProperties(fileConfig.Properties as IItem);
            properties.__metadata = { type: "SP.ListItem" };
            file.listItemAllFields.update(properties)
                .then((itemUpdateResult) => {
                    Util.Resolve<Item>(resolve, this.handlerName, `Updated item: '${fileConfig.Name}'`, itemUpdateResult.item);
                })
                .catch((error) => {
                    Util.Reject<void>(reject, this.handlerName, `Error while updating item with title '${fileConfig.Name}': ` + Util.getErrorMessage(error));
                });
        });
    }

    private resolvePreTask(fileConfig: IFile): Promise<any> {
        let promise = Promise.resolve();

        if (fileConfig.DataConnections) {
            promise = this.updateDataConnection(fileConfig);
        }

        return promise;
    }

    private createProperties(itemConfig: IItem) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(itemConfig);
        let parsedObject = JSON.parse(stringifiedObject);

        delete parsedObject.ControlOption;
        delete parsedObject.Name;
        delete parsedObject.DataConnections;

        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }

    //#region "Pre Tasks"
    private updateDataConnection(fileConfig: IFile): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let connections = [];
            let context = SP.ClientContext.get_current();
            let lists = context.get_web().get_lists();
            for (let connConf of fileConfig.DataConnections) {
                let connection = {
                    ListName: connConf.List,
                    ListId: lists.getByTitle(connConf.List),
                    ListRootFolderUrl: lists.getByTitle(connConf.List).get_rootFolder(),
                    ViewName: connConf.View,
                    ViewId: lists.getByTitle(connConf.List).get_views().getByTitle(connConf.View),
                    WebUrl: window._spPageContextInfo.webAbsoluteUrl,
                };
                connections.push(connection);
                context.load(connection.ListId);
                context.load(connection.ListRootFolderUrl);
                context.load(connection.ViewId);
            }
            context.executeQueryAsync(
                (sender, args) => {
                    for (let connection of connections) {
                        connection.ListId = connection.ListId.get_id().toString("B");
                        connection.ListRootFolderUrl = connection.ListRootFolderUrl.get_serverRelativeUrl();
                        connection.ViewId = connection.ViewId.get_id().toString("B");
                    }
                    fs.writeFileSync(`${fileConfig.Src}.json`, JSON.stringify(connections));
                    let ps = spawn("powershell.exe",
                        [
                            ".\\updateDataConnection.ps1",
                            "-File",
                            fileConfig.Src,
                            "-ConnectionTemplate",
                            fileConfig.DataConnectionTemplate,
                        ]);
                    ps.stdout.on("data", (data) => {
                        Logger.write(`${this.handlerName} - Processing data connection for file '${fileConfig.Name}': '${data}'`, Logger.LogLevel.Info);
                    });
                    ps.stderr.on("data", (data) => {
                        Util.Reject<void>(reject, this.handlerName,
                            `Error while updating data connection for '${fileConfig.Name}':  ${data}`);
                    });
                    ps.on("exit", () => {
                        fs.unlinkSync(`${fileConfig.Src}.json`);
                        resolve();
                    });
                    ps.stdin.end();
                },
                (sender, args) => {
                    Util.Reject<void>(reject, this.handlerName, `Error while updating data connection for '${fileConfig.Name}': `
                        + `${Util.getErrorMessageFromQuery(args.get_message(), args.get_stackTrace())}`);
                }
            );
        });
    }
    //#endregion "Pre Tasks"
}
