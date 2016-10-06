import { IField } from "../Types/IField";
import { IList } from "../Types/IList";

export interface SiteConfig {
    Url: string;
    LayoutsHive: string;
    Fields: IField[];
    Lists: IList[];
}
