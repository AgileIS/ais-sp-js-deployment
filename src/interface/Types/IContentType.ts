export interface IContentType {
    Id: string;
    Name: string;
    Description: string;
    DisplayFormUrl: string;
    DisplayFormTemplateName: string;
    EditFormUrl: string;
    EditFormTemplateName: string;
    NewFormUrl: string;
    NewFormTemplateName: string;
    DocumentTemplate: string;
    Group: string;
    Hidden: boolean;
    JSLink: string;
    Sealed: boolean;
    ReadOnly: boolean;
    FieldLinks: Array<string>
    ControlOption: string;
}
