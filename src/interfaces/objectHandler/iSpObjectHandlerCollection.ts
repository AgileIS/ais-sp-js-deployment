import { ISPObjectHandler } from "./iSpObjectHandler";

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
        Solutions: ISPObjectHandler;
}
