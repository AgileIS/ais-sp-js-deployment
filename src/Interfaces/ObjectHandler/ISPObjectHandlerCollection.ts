import { ISPObjectHandler } from "./ISPObjectHandler";

export interface ISPObjectHandlerCollection {
        Features: ISPObjectHandler;
        Sites: ISPObjectHandler;
        ContentTypes: ISPObjectHandler;
        Lists: ISPObjectHandler;
        Fields: ISPObjectHandler;
        Views: ISPObjectHandler;
        Items: ISPObjectHandler;
        Navigation: ISPObjectHandler;
        Files: ISPObjectHandler;
}
