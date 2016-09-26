import { INavigationNode } from "./inavigationnode";

export interface INavigation {
    UseShared: boolean;
    ReCreateQuicklaunch: boolean;
    QuickLaunch: Array<INavigationNode>;
}
