import { IFolder } from "./IFolder";
import { IListFieldRef } from "./IListFieldRef";
import { IField } from "./IField";
import { IView } from "./IView";
import { ISecurity } from "./ISecurity";
import { IContentTypeBinding } from "./IContentTypeBinding";

export interface IList {
    Title: string;
    Url: string;
    Description: string;
    DocumentTemplate: string;
    OnQuickLaunch: boolean;
    TemplateType: number;
    EnableVersioning: boolean;
    EnableMinorVersions: boolean;
    EnableModeration: boolean;
    EnableFolderCreation: boolean;
    EnableAttachments: boolean;
    RemoveExistingContentTypes: boolean;
    RemoveExistingViews: boolean;
    NoCrawl: boolean;
    DefaultDisplayFormUrl: string;
    DefaultEditFormUrl: string;
    DefaultNewFormUrl: string;
    DraftVersionVisibility: string;
    ImageUrl: string;
    Hidden: boolean;
    ForceCheckout: boolean;
    ContentTypeBindings: Array<IContentTypeBinding>;
    FieldRefs: Array<IListFieldRef>;
    Fields: Array<IField>;
    Folders: Array<IFolder>;
    Views: Array<IView>;
    DataRows: Array<Object>;
    Security: ISecurity;
}