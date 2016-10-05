import { IField } from "../Types/IField";
import { IList } from "../Types/IList";

export interface SiteCollectionConfig {
    Url: string;
    LayoutsHive: string;
    Fields: IField[];
    Lists: IList[];
}
