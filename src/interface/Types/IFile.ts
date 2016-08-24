import { IWebPart } from "./IWebPart";
import { IHiddenView } from "./IHiddenView";

export interface IFile {
    Overwrite: boolean;
    Dest: string;
    Src: string;
    Properties: Object;
    RemoveExistingWebParts: boolean;
    WebParts: Array<IWebPart>;
    Views: Array<IHiddenView>;
}