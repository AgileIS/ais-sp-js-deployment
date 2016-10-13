import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ContentType } from "@agileis/sp-pnp-js/lib/sharepoint/rest/contenttypes";
import { ISPObjectHandler } from "../interfaces/objectHandler/iSpObjectHandler";
import { IContentType } from "../interfaces/types/iContentType";
import { IPromiseResult } from "../interfaces/iPromiseResult";
import { PromiseResult } from "../promiseResult";
import { ControlOption } from "../constants/controlOption";
import { Util } from "../util/util";

export class ContentTypeHandler implements ISPObjectHandler {
    public execute(contentTypeConfig: IContentType, parentPromise: Promise<IPromiseResult<Web>>): Promise<IPromiseResult<void | ContentType>> {
        return new Promise<IPromiseResult<ContentType>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, contentTypeConfig.Id,
                        `Content type handler parent promise value result is null or undefined for the content type with the id '${contentTypeConfig.Id}'!`);
                } else {
                    if (contentTypeConfig && contentTypeConfig.Id && contentTypeConfig.Name) {
                        let web = promiseResult.value;
                        let context = new SP.ClientContext(_spPageContextInfo.webAbsoluteUrl);
                        Util.tryToProcess(contentTypeConfig.Id, () => {
                            context = new SP.ClientContext(_spPageContextInfo.webAbsoluteUrl);
                            return this.processingContentTypeConfig(contentTypeConfig, context); })
                            .then((contentTypeProcessingResult) => {
                                let resolveValue = undefined;
                                if (contentTypeProcessingResult.value) {
                                    let contentType = (<SP.ContentType>contentTypeProcessingResult.value);
                                    resolveValue = web.contentTypes.getById(contentType.get_id().get_stringValue());
                                }
                                resolve(new PromiseResult(contentTypeProcessingResult.message, resolveValue));
                            })
                            .catch((error) => { reject(error); });
                    } else {
                        Util.Reject<void>(reject, "Unknown content type", `Error while processing content type: Content type id or/and name are undefined.`);
                    }
                }
            });
        });
    }

    private processingContentTypeConfig(contentTypeConfig: IContentType, clientContext: SP.ClientContext): Promise<IPromiseResult<void | SP.ContentType>> {
        return new Promise<IPromiseResult<void | SP.ContentType>>((resolve, reject) => {
            let processingText = contentTypeConfig.ControlOption === ControlOption.ADD || contentTypeConfig.ControlOption === undefined || contentTypeConfig.ControlOption === ""
                ? "Add" : contentTypeConfig.ControlOption;
            Logger.write(`Processing ${processingText} content type: '${contentTypeConfig.Id}'.`, Logger.LogLevel.Info);

            let web = clientContext.get_web();
            let rootWeb = clientContext.get_site().get_rootWeb();
            this.getContentType(contentTypeConfig, clientContext, web)
                .then(contentTypeResult => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | SP.ContentType>> = undefined;
                    if (contentTypeResult) {
                        switch (contentTypeConfig.ControlOption) {
                            case ControlOption.UPDATE:
                                processingPromise = this.updateContentType(contentTypeConfig, contentTypeResult, web);
                                break;
                            case ControlOption.DELETE:
                                processingPromise = this.deleteContentType(contentTypeConfig, contentTypeResult);
                                break;
                            default:
                                Util.Resolve<SP.ContentType>(resolve, contentTypeConfig.Id,
                                    `Content type with the id '${contentTypeConfig.Id}' does not have to be added, because it already exists.`, contentTypeResult);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (contentTypeConfig.ControlOption) {
                            case ControlOption.DELETE:
                                Util.Resolve<void>(resolve, contentTypeConfig.Id, `Content type with the id '${contentTypeConfig.Id}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.UPDATE:
                                contentTypeConfig.ControlOption = ControlOption.ADD;
                            default:
                                processingPromise = this.addContentType(contentTypeConfig, rootWeb.get_contentTypes(), web);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((contentTypeProsssingResult) => { resolve(contentTypeProsssingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write(`Content type handler processing promise is undefined for the content type with the id '${contentTypeConfig.Id}'!`, Logger.LogLevel.Error);
                    }
                })
                .catch(error => { Util.Reject<void>(reject, contentTypeConfig.Id, `Error while requesting content type with the id '${contentTypeConfig.Id}': ${Util.getErrorMessage(error)}`); });
        });
    }

    private setContentTypeProperties(contentTypeConfig: IContentType, contentType: SP.ContentType): void {
        for (let prop in contentTypeConfig) {
            switch (prop) {
                case "Description":
                    contentType.set_description(contentTypeConfig.Description);
                    break;
                case "DisplayFormTemplateName":
                    contentType.set_displayFormTemplateName(contentTypeConfig.DisplayFormTemplateName);
                    break;
                case "DisplayFormUrl":
                    contentType.set_displayFormUrl(contentTypeConfig.DisplayFormUrl);
                    break;
                case "DocumentTemplate":
                    contentType.set_documentTemplate(contentTypeConfig.DocumentTemplate);
                    break;
                case "EditFormTemplateName":
                    contentType.set_editFormTemplateName(contentTypeConfig.EditFormTemplateName);
                    break;
                case "EditFormUrl":
                    contentType.set_editFormUrl(contentTypeConfig.EditFormUrl);
                    break;
                case "Group":
                    contentType.set_group(contentTypeConfig.Group);
                    break;
                case "Hidden":
                    contentType.set_hidden(contentTypeConfig.Hidden);
                    break;
                case "JSLink":
                    contentType.set_jsLink(contentTypeConfig.JSLink);
                    break;
                case "Name":
                    contentType.set_name(contentTypeConfig.Name);
                    break;
                case "NewFormTemplateName":
                    contentType.set_newFormTemplateName(contentTypeConfig.NewFormTemplateName);
                    break;
                case "NewFormUrl":
                    contentType.set_newFormUrl(contentTypeConfig.NewFormUrl);
                    break;
                case "ReadOnly":
                    contentType.set_readOnly(contentTypeConfig.ReadOnly);
                    break;
                case "Sealed":
                    contentType.set_sealed(contentTypeConfig.Sealed);
                    break;
            }
        }
    }

    private setContentTypeFieldLinks(contentTypeConfig: IContentType, contentType: SP.ContentType, web: SP.Web) {
        if (contentTypeConfig.FieldLinks && contentTypeConfig.FieldLinks.length > 0) {
            let currentFieldLinksInternalNames: Array<string> = new Array<string>();
            let existingFieldLinks: { [internalName: string]: SP.Guid } = {};
            let fieldLinks = contentType.get_fieldLinks();

            let existingFieldLinksEnumerator = fieldLinks.getEnumerator();
            while (existingFieldLinksEnumerator.moveNext()) {
                let fieldlink = existingFieldLinksEnumerator.get_current();
                existingFieldLinks[fieldlink.get_name()] = fieldlink.get_id();
            }

            contentTypeConfig.FieldLinks.forEach((fieldLink, index, array) => {
                let spfieldLink: SP.FieldLink = undefined;

                if (existingFieldLinks[fieldLink.InternalName]) {
                    spfieldLink = fieldLinks.getById(existingFieldLinks[fieldLink.InternalName]);
                    delete existingFieldLinks[fieldLink.InternalName];
                } else {
                    let fieldLinkCreationInfo = new SP.FieldLinkCreationInformation();
                    let field = web.get_availableFields().getByInternalNameOrTitle(fieldLink.InternalName);
                    fieldLinkCreationInfo.set_field(field);
                    spfieldLink = fieldLinks.add(fieldLinkCreationInfo);
                }

                if (fieldLink.Required) {
                    spfieldLink.set_required(fieldLink.Required === true ? fieldLink.Required : false);
                }

                if (fieldLink.Hidden) {
                    spfieldLink.set_hidden(fieldLink.Hidden === true ? fieldLink.Hidden : false);
                }

                currentFieldLinksInternalNames.push(fieldLink.InternalName);
            });

            for (let bwFieldLink in existingFieldLinks) {
                fieldLinks.getById(existingFieldLinks[bwFieldLink]).deleteObject();
            }

            fieldLinks.reorder(currentFieldLinksInternalNames);
        }
    }

    private getContentTypeCreationInfo(contentTypeConfig: IContentType): SP.ContentTypeCreationInformation {
        let contentTypeCreationInfo = new SP.ContentTypeCreationInformation();
        contentTypeCreationInfo.set_id(contentTypeConfig.Id);
        contentTypeCreationInfo.set_name(contentTypeConfig.Name);
        if (contentTypeConfig.Description) {
            contentTypeCreationInfo.set_description(contentTypeConfig.Description);
        }
        if (contentTypeConfig.Group) {
            contentTypeCreationInfo.set_group(contentTypeConfig.Group);
        }
        return contentTypeCreationInfo;
    }

    private addContentType(contentTypeConfig: IContentType, contentTypeCollection: SP.ContentTypeCollection, web: SP.Web): Promise<IPromiseResult<SP.ContentType>> {
        return new Promise<IPromiseResult<SP.ContentType>>((resolve, reject) => {
            let contentTypeCreationInfo = this.getContentTypeCreationInfo(contentTypeConfig);
            let newContentType = contentTypeCollection.add(contentTypeCreationInfo);
            let context = contentTypeCollection.get_context();
            context.load(newContentType, "Name", "Id", "FieldLinks");
            context.executeQueryAsync(
                (sender, args) => {
                    this.updateContentType(contentTypeConfig, newContentType, web)
                        .then((contentTypeUpdateResult) => {
                            Util.Resolve<SP.ContentType>(resolve, contentTypeConfig.Id, `Created and updated content Type: '${contentTypeConfig.Id}'.`, contentTypeUpdateResult.value);
                        })
                        .catch((error) => {
                            this.tryToDeleteCorruptedContentType(contentTypeConfig, context, web)
                                .then(() => {
                                    Util.Reject<void>(reject, contentTypeConfig.Id,
                                        `Error while adding ContentType with the id '${contentTypeConfig.Id}' - corrupted ContentType deleted`);
                                })
                                .catch(() => {
                                    Util.Reject<void>(reject, contentTypeConfig.Id,
                                        `Error while adding ContentType with the id '${contentTypeConfig.Id}' - corrupted ContentType not deleted`);
                                });
                        });
                },
                (sender, args) => {
                    Util.Reject<void>(reject, contentTypeConfig.Id, `Error while adding content type with the id '${contentTypeConfig.Id}': `
                        + `${Util.getErrorMessageFromQuery(args.get_message(), args.get_stackTrace())}`);
                }
            );
        });
    }

    private updateContentType(contentTypeConfig: IContentType, contentType: SP.ContentType, web: SP.Web): Promise<IPromiseResult<SP.ContentType>> {
        return new Promise<IPromiseResult<SP.ContentType>>((resolve, reject) => {
            this.setContentTypeProperties(contentTypeConfig, contentType);
            this.setContentTypeFieldLinks(contentTypeConfig, contentType, web);
            contentType.update(true);
            contentType.get_context().executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<SP.ContentType>(resolve, contentTypeConfig.Id, `Updated content type: '${contentTypeConfig.Id}'.`, contentType);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, contentTypeConfig.Id, `Error while updating content type with the id '${contentTypeConfig.Id}': `
                        + `${Util.getErrorMessageFromQuery(args.get_message(), args.get_stackTrace())}`);
                }
            );
        });
    }

    private deleteContentType(contentTypeConfig: IContentType, contentType: SP.ContentType): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            contentType.deleteObject();
            contentType.get_context().executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, contentTypeConfig.Id, `Deleted content type: '${contentTypeConfig.Id}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, contentTypeConfig.Id, `Error while deleting content type with the id '${contentTypeConfig.Id}': `
                        + `${Util.getErrorMessageFromQuery(args.get_message(), args.get_stackTrace())}`);
                }
            );
        });
    }

    private tryToDeleteCorruptedContentType(contentTypeConfig: IContentType, clientContext: SP.ClientRuntimeContext, web: SP.Web): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            Logger.write(`Try to delete corrupted ContentType with id: '${contentTypeConfig.Id}'`);
            this.getContentType(contentTypeConfig, clientContext, web)
                .then((contentTypeResult) => {
                    if (contentTypeResult) {
                        this.deleteContentType(contentTypeConfig, contentTypeResult)
                            .then(() => { resolve(); })
                            .catch(error => { reject(error); });
                    } else { reject(`No ContentType with ID '${contentTypeConfig.Id}' found`); }
                })
                .catch((error) => { reject(error); });
        });
    }

    private getContentType(contentTypeConfig: IContentType, clientContext: SP.ClientRuntimeContext, web: SP.Web): Promise<SP.ContentType> {
        return new Promise<SP.ContentType | void>((resolve, reject) => {
            let webContentType = web.get_contentTypes().getById(contentTypeConfig.Id);
            let siteContentType = web.get_availableContentTypes().getById(contentTypeConfig.Id);
            let contentType: SP.ContentType = undefined;
            clientContext.load(webContentType, "Id", "Name", "FieldLinks");
            clientContext.load(siteContentType, "Id", "Name", "FieldLinks");
            clientContext.executeQueryAsync(
                (sender, args) => {
                    if (!siteContentType.get_serverObjectIsNull() || !webContentType.get_serverObjectIsNull()) {
                        Logger.write(`Found ContentType with id: '${contentTypeConfig.Id}'`);
                        if (!webContentType.get_serverObjectIsNull()) {
                            contentType = webContentType;
                        } else {
                            contentType = siteContentType;
                        }
                    }
                    resolve(contentType);
                },
                (sender, args) => {
                    reject(Util.getErrorMessageFromQuery(args.get_message(), args.get_stackTrace()));
                });
        });

    }
}
