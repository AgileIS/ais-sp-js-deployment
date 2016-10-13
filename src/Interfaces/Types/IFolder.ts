import { IFile } from "./iFile";

export interface IFolder {
    Name: string;
    DefaultValues: Object;
    Files: Array<IFile | IFolder>;
    ControlOption: string;
}
