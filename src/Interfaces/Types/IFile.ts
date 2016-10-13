import { IWebPart } from "./iWebPart";
import { IHiddenView } from "./iHiddenView";
import { IDataConnection } from "./iDataConnection";

export interface IFile {
    Overwrite: boolean;
    Dest: string;
    Src: string;
    Properties: Object;
    RemoveExistingWebParts: boolean;
    WebParts: Array<IWebPart>;
    Views: Array<IHiddenView>;
    Name: string;
    ControlOption: string;
    DataConnections: Array<IDataConnection>;
    DataConnectionTemplate: string;
}
