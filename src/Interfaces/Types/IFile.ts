import { IWebPart } from "./IWebPart";
import { IHiddenView } from "./IHiddenView";
import { IDataConnection } from "./IDataConnection";

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
