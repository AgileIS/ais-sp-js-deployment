import { IFile } from "./ifile";

export interface IFolder {
    Name: string;
    DefaultValues: Object;
    Files: Array<IFile | IFolder>;
    ControlOption: string;
}
