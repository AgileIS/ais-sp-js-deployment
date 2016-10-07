import { IList } from "./IList";
import { ICustomAction } from "./ICustomaction";
import { IFeature } from "./IFeature";
import { IFile } from "./IFile";
import { IField } from "./IField";
import { INavigation } from "./INavigation";
import { IComposedLook } from "./IComposedlook";
import { IWebSettings } from "./IWebsettings";
import { IContentType } from "./IContentType";
import { IPropertyBagEntry } from "./IPropertyBagEntry";

export interface ISite {
    Url: string;
    ContentTypes: Array<IContentType>;
    Lists: Array<IList>;
    Files: Array<IFile>;
    Fields: Array<IField>;
    Navigation: INavigation;
    CustomActions: Array<ICustomAction>;
    ComposedLook: IComposedLook;
    PropertyBagEntries: Array<IPropertyBagEntry>;
    Parameters: Object;
    WebSettings: IWebSettings;
    Features: Array<IFeature>;
    ControlOption: string;
    Title: string;
    Description: string;
    EnableMinimalDownload: boolean;
    QuickLaunchEnabled: boolean;
    ServerRelativeUrl: string;
    TreeViewEnabled: boolean;
    RecycleBinEnabled: boolean;
    LayoutsHive: string;
    Template: string;
    Language: number;
    InheritPermissions: boolean;
    Identifier: string;
}
