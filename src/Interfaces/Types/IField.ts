export interface IField {
    ShowInDisplayForm: boolean;
    ShowInEditForm: boolean;
    ShowInNewForm: boolean;
    CanBeDeleted: boolean;
    DefaultValue: string;
    Description: string;
    EnforceUniqueValues: boolean;
    Direction: string;
    EntityPropertyName: string;
    FieldType: string;
    Filterable: boolean;
    Group: string;
    Hidden: boolean;
    ID: string;
    Indexed: boolean;
    InternalName: string;
    JsLink: string;
    ReadOnlyField: boolean;
    Required: boolean;
    SchemaXml: string;
    StaticName: string;
    Title: string;
    TypeAsString: string;
    TypeDisplayName: string;
    TypeShortDescription: string;
    ValidationFormula: string;
    ValidationMessage: string;
    Type: string;
    Formula: string;
    DateFormat: string;
    OutputType: string;
    ControlOption: string;
    LookupList: string;
    LookupField: string;
    Multivalue: boolean;
    RelationshipDeleteBehavior: string;
    DependendFields: Array<IFieldDependendLookup>;
}

export interface IFieldDependendLookup {
    InternalName: string;
    Title: string;
}
