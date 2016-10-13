import { INavigationNode } from "./iNavigationNode";

export interface INavigation {
    UseShared: boolean;
    ReCreateQuicklaunch: boolean;
    QuickLaunch: Array<INavigationNode>;
}
