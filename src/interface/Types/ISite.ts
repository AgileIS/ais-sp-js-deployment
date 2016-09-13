import { IList } from "./IList";
import { ICustomAction } from "./ICustomaction";
import { IFeature } from "./IFeature";
import { IFile } from "./IFile";
import { IField } from "./IField";
import { INavigation } from "./INavigation";
import { IComposedLook } from "./IComposedlook";
import { IWebSettings } from "./IWebsettings";
import { IContentType } from "./IContentType"

export interface ISite {
    Url: string;
    ContentTypes: Array<IContentType>;
    Lists: Array<IList>;
    Files: Array<IFile>;
    Fields: Array<IField>;
    Navigation: INavigation;
    CustomActions: Array<ICustomAction>;
    ComposedLook: IComposedLook;
    PropertyBagEntries: Object;
    Parameters: Object;
    WebSettings: IWebSettings;
    Features: Array<IFeature>;
}