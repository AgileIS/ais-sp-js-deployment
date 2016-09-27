export interface INavigationNode {
    Title: string;
    Url: string;
    Children: Array<INavigationNode>;
    IsExternal: boolean;
    ControlOption: string;
}
