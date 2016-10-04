import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ContentType } from "@agileis/sp-pnp-js/lib/sharepoint/rest/contenttypes";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IContentType } from "../Interfaces/Types/IContentType";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { PromiseResult } from "../PromiseResult";
import { ControlOption } from "../Constants/ControlOption";
import { Util } from "../Util/Util";

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
                        let context = SP.ClientContext.get_current();
                        this.processingContentTypeConfig(contentTypeConfig, context)
                            .then((contentTypeProsssingResult) => {
                                let resolveValue = undefined;
                                if (contentTypeProsssingResult.value) {
                                    let contentType = (<SP.ContentType>contentTypeProsssingResult.value);
                                    resolveValue = web.contentTypes.getById(contentType.get_id().get_stringValue());
                                }
                                resolve(new PromiseResult(contentTypeProsssingResult.message, resolveValue));
                            })
                            .catch((error) => {
                                Util.Retry(error, contentTypeConfig.Id,
                                    () => {
                                        return this.processingContentTypeConfig(contentTypeConfig, context);
                                    });
                            });
                    } else {
                        Util.Reject<void>(reject, contentTypeConfig.Id, `Error while processing content type with the id '${contentTypeConfig.Id}': Content type id or/and name are undefined.`);
                    }
                }
            });
        });
    }

    private processingContentTypeConfig(contentTypeConfig: IContentType, clientContext: SP.ClientContext): Promise<IPromiseResult<void | SP.ContentType>> {
        return new Promise<IPromiseResult<void | SP.ContentType>>((resolve, reject) => {
            let processingText = contentTypeConfig.ControlOption === ControlOption.Add || contentTypeConfig.ControlOption === undefined || contentTypeConfig.ControlOption === ""
                ? "Add" : contentTypeConfig.ControlOption;
            Logger.write(`Processing ${processingText} content type: '${contentTypeConfig.Id}'.`, Logger.LogLevel.Info);

            let web = clientContext.get_web();
            let rootWeb = clientContext.get_site().get_rootWeb();
            let webContentType = web.get_contentTypes().getById(contentTypeConfig.Id);
            let siteContentType = web.get_availableContentTypes().getById(contentTypeConfig.Id);
            clientContext.load(webContentType, "Id", "Name", "FieldLinks");
            clientContext.load(siteContentType, "Id", "Name", "FieldLinks");
            clientContext.executeQueryAsync(
                (sender, args) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | SP.ContentType>> = undefined;

                    if (!siteContentType.get_serverObjectIsNull() || !webContentType.get_serverObjectIsNull()) {
                        Logger.write(`Found ContentType with id: '${contentTypeConfig.Id}'`);
                        let contentType: SP.ContentType = undefined;
                        if (!webContentType.get_serverObjectIsNull()) {
                            contentType = webContentType;
                        } else {
                            contentType = siteContentType;
                        }

                        switch (contentTypeConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateContentType(contentTypeConfig, contentType, web);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteContentType(contentTypeConfig, contentType);
                                break;
                            default:
                                Util.Resolve<SP.ContentType>(resolve, contentTypeConfig.Id,
                                    `Content type with the id '${contentTypeConfig.Id}' does not have to be added, because it already exists.`, contentType);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (contentTypeConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Resolve<void>(resolve, contentTypeConfig.Id, `Content type with the id '${contentTypeConfig.Id}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.Update:
                                contentTypeConfig.ControlOption = ControlOption.Add;
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
                },
                (sender, args) => {
                    Util.Reject<void>(reject, contentTypeConfig.Id,
                        `Error while requesting content type with the id '${contentTypeConfig.Id}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                });
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
            let fieldLinks = contentType.get_fieldLinks();

            this.updateTitleFieldLink(contentTypeConfig, fieldLinks);
            contentTypeConfig.FieldLinks.forEach((fieldLink, index, array) => {
                let fieldLinkCreationInfo = new SP.FieldLinkCreationInformation();
                let field = web.get_availableFields().getByInternalNameOrTitle(fieldLink.InternalName);
                fieldLinkCreationInfo.set_field(field);
                let spfieldLink = fieldLinks.add(fieldLinkCreationInfo);

                if (fieldLink.Required) {
                    spfieldLink.set_required(fieldLink.Required === true ? fieldLink.Required : false);
                }

                if (fieldLink.Hidden) {
                    spfieldLink.set_hidden(fieldLink.Hidden === true ? fieldLink.Hidden : false);
                }

                currentFieldLinksInternalNames.push(fieldLink.InternalName);
            });

            fieldLinks.reorder(currentFieldLinksInternalNames);
        }

    }

    private updateTitleFieldLink(contentTypeConfig: IContentType, fieldLinks: SP.FieldLinkCollection): void {
        let item = contentTypeConfig.FieldLinks.filter( (fLink, fIndex) => { return fLink.InternalName === "Title"; });
            contentTypeConfig.FieldLinks.splice(contentTypeConfig.FieldLinks.indexOf(item[0]), 1);
            let e = fieldLinks.getEnumerator();
            while (e.moveNext()) {
                let current = e.get_current();
                if (current.get_name() === "Title") {
                    current.set_required(item[0] ? (item[0].Required ? item[0].Required : false) : false);
                    current.set_hidden(item[0] ? (item[0].Hidden ? item[0].Hidden : false) : true);
                }
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
                            Util.Resolve<SP.ContentType>(resolve, contentTypeConfig.Id,
                                `Created and updated content Type: '${contentTypeConfig.Id}'.`,
                                contentTypeUpdateResult.value);
                        })
                        .catch((error) => { reject(error); });
                },
                (sender, args) => {
                    Util.Reject<void>(reject, contentTypeConfig.Id,
                        `Error while adding content type with the id '${contentTypeConfig.Id}':  ${args.get_message()} '\n' ${args.get_stackTrace()}`);
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
                    Util.Reject<void>(reject, contentTypeConfig.Id,
                        `Error while updating content type with the id '${contentTypeConfig.Id}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
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
                    Util.Reject<void>(reject, contentTypeConfig.Id,
                        `Error while deleting content type with the id '${contentTypeConfig.Id}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                }
            );
        });
    }
}
