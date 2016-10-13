import { IField } from "../types/iField";
import { IList } from "../types/iList";

export interface ISiteConfig {
    Url: string;
    LayoutsHive: string;
    Fields: IField[];
    Lists: IList[];
}
