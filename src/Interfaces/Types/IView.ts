import { IViewField } from "../types/iViewField";

export interface IView {
    Title: string;
    NewTitle: string;
    InternalName: string;
    Paged: boolean;
    PersonalView: boolean;
    ViewQuery: string;
    RowLimit: number;
    Scope: number;
    SetAsDefaultView: boolean;
    ViewFields: Array<IViewField>;
    ViewTypeKind: string;
    ControlOption: string;
    ParentListId: any;
}
