import { IField } from "../Types/IField";
import { IList } from "../Types/IList";

export interface ISiteConfig {
    Url: string;
    LayoutsHive: string;
    Fields: IField[];
    Lists: IList[];
}
