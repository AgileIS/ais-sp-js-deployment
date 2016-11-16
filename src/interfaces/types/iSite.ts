import { IList } from "./iList";
import { ICustomAction } from "./iCustomAction";
import { IFeature } from "./iFeature";
import { IFile } from "./iFile";
import { IField } from "./iField";
import { INavigation } from "./iNavigation";
import { IComposedLook } from "./iComposedLook";
import { IWebSettings } from "./iWebSettings";
import { IContentType } from "./iContentType";
import { IPropertyBagEntry } from "./iPropertyBagEntry";

export interface ISite {
    Url: string;
    WebApplicationUrl: string;
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
