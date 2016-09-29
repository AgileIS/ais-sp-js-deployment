import { IViewField } from "../Types/IViewField";

export interface IView {
    Title: string;
    Paged: boolean;
    PersonalView: boolean;
    Query: string;
    RowLimit: number;
    Scope: number;
    SetAsDefaultView: boolean;
    ViewFields: Array<IViewField>;
    ViewTypeKind: string;
    ControlOption: string;
    ParentListId: any;
}
