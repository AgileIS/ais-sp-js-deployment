import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ContentType } from "@agileis/sp-pnp-js/lib/sharepoint/rest/ContentTypes";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IContentType } from "../interface/Types/IContentType";
import { ControlOption } from "../Constants/ControlOption";
import { Reject, Resolve } from "../Util/Util";

export class ContentTypeHandler implements ISPObjectHandler {
    public execute(contentTypeConfig: IContentType, parentPromise: Promise<Web>): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            parentPromise.then((parentWeb) => {
                if (contentTypeConfig.Id && contentTypeConfig.Name) {
                    this.processingContentTypeConfig(contentTypeConfig, parentWeb)
                        .then((contentTypeProsssingResult) => { resolve(contentTypeProsssingResult); })
                        .catch((error) => { reject(error); });
                } else {
                    Reject(reject, `Error while processing content type with the name '${contentTypeConfig.Name}': Content type id or/and name are undefined`, contentTypeConfig.Name);
                }
            });
        });
    }

    private processingContentTypeConfig(contentTypeConfig: IContentType, parentWeb: Web): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            let processingText = contentTypeConfig.ControlOption === ControlOption.Add || contentTypeConfig.ControlOption === undefined || contentTypeConfig.ControlOption === ""
                ? "Add" : contentTypeConfig.ControlOption;
            Logger.write(`Processing ${processingText} content type: '${contentTypeConfig.Name}'`, Logger.LogLevel.Info);

            let context = new SP.ClientContext(parentWeb.toUrl().split("/_")[0]);
            let web = context.get_web();
            let rootWeb = context.get_site().get_rootWeb();
            let webContentType = web.get_contentTypes().getById(contentTypeConfig.Id);
            let siteContentType = web.get_availableContentTypes().getById(contentTypeConfig.Id);
            context.load(webContentType, "Id", "Name");
            context.load(siteContentType, "Id", "Name");
            context.executeQueryAsync(
                (sender, args) => {
                    let processingPromise: Promise<SP.ContentType> = undefined;

                    if (!siteContentType.get_serverObjectIsNull() || !webContentType.get_serverObjectIsNull()) {
                        let contentType: SP.ContentType = undefined;
                        if (!webContentType.get_serverObjectIsNull()) {
                            contentType = webContentType;
                        } else {
                            contentType = siteContentType;
                        }

                        let restContentType = parentWeb.contentTypes.getById(contentType.get_id().get_typeId());

                        switch (contentTypeConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateContentType(contentTypeConfig, contentType, web);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteContentType(contentTypeConfig, contentType);
                                break;
                            default:
                                Resolve(resolve, `Content type with the name '${contentTypeConfig.Name}' already exists in target web`, contentTypeConfig.Name, restContentType);
                                break;
                        }
                    } else {
                        switch (contentTypeConfig.ControlOption) {
                            case ControlOption.Update:
                            case ControlOption.Delete:
                                Reject(reject, `Content type with the name '${contentTypeConfig.Name}' does not exists`, contentTypeConfig.Name);
                                break;
                            default:
                                processingPromise = this.addContentType(contentTypeConfig, rootWeb.get_contentTypes(), web);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((contentTypeProsssingResult) => {
                                let resolveResult = undefined;
                                if (typeof contentTypeProsssingResult !== "string") {
                                    resolveResult = parentWeb.contentTypes.getById(contentTypeProsssingResult.get_id().get_stringValue());
                                }
                                resolve(resolveResult);
                            })
                            .catch((error) => { reject(error); });
                    } else {
                        Logger.write("Content type handler processing promise is undefined!");
                    }
                },
                (sender, args) => {
                    Reject(reject, `Error while requesting content type with the name '${contentTypeConfig.Name}': ${args.get_message()} '\n' ${args.get_stackTrace()}`, contentTypeConfig.Name);
                });
        });
    }

    private setContentTypeProperties(contentTypeConfig: IContentType, contentType: SP.ContentType, parentWeb: SP.Web): void {
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

        if (contentTypeConfig.FieldLinks && contentTypeConfig.FieldLinks.length > 0) {
            let fieldLinks = contentType.get_fieldLinks();
            contentTypeConfig.FieldLinks.forEach((fieldInternalName, index, array) => {
                let fieldLink = new SP.FieldLinkCreationInformation();
                let field = parentWeb.get_availableFields().getByInternalNameOrTitle(fieldInternalName);
                fieldLink.set_field(field);
                fieldLinks.add(fieldLink);
            });
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


    private addContentType(contentTypeConfig: IContentType, contentTypeCollection: SP.ContentTypeCollection, parentWeb: SP.Web): Promise<SP.ContentType> {
        return new Promise<SP.ContentType>((resolve, reject) => {
            let contentTypeCreationInfo = this.getContentTypeCreationInfo(contentTypeConfig);
            let newContentType = contentTypeCollection.add(contentTypeCreationInfo);
            let context = contentTypeCollection.get_context();
            context.load(newContentType, "Name", "Id", "FieldLinks");
            context.executeQueryAsync(
                (sender, args) => {
                    this.updateContentType(contentTypeConfig, newContentType, parentWeb)
                        .then((contentTypeUpdateResult) => {
                            Resolve(resolve, `Created and updated content Type: '${contentTypeConfig.Name}'.`,
                                contentTypeConfig.Name, contentTypeUpdateResult);
                        })
                        .catch((error) => { reject(error); });
                },
                (sender, args) => {
                    Reject(reject, `Error while adding content type with the name '${contentTypeConfig.Name}':  ${args.get_message()} '\n' ${args.get_stackTrace()}`, contentTypeConfig.Name);
                }
            );
        });
    }

    private updateContentType(contentTypeConfig: IContentType, contentType: SP.ContentType, parentWeb: SP.Web): Promise<SP.ContentType> {
        return new Promise<SP.ContentType>((resolve, reject) => {
            this.setContentTypeProperties(contentTypeConfig, contentType, parentWeb);
            contentType.update(true);
            contentType.get_context().executeQueryAsync(
                (sender, args) => {
                    Resolve(resolve, `Updated content type: '${contentTypeConfig.Name}'`, contentTypeConfig.Name, contentType);
                },
                (sender, args) => {
                    Reject(reject, `Error while updating content type with the name '${contentTypeConfig.Name}': ${args.get_message()} '\n' ${args.get_stackTrace()}`, contentTypeConfig.Name);
                }
            );
        });
    }

    private deleteContentType(contentTypeConfig: IContentType, contentType: SP.ContentType): Promise<SP.ContentType> {
        return new Promise<SP.ContentType>((resolve, reject) => {
            contentType.deleteObject();
            contentType.get_context().executeQueryAsync(
                (sender, args) => {
                    Resolve(resolve, `Deleted content Type: '${contentTypeConfig.Name}'`, contentTypeConfig.Name);
                },
                (sender, args) => {
                    Reject(reject, `Error while deleting content type with the name '${contentTypeConfig.Name}': ${args.get_message()} '\n' ${args.get_stackTrace()}`, contentTypeConfig.Name);
                }
            );
        });
    }
}