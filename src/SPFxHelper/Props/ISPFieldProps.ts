/**
 * Interface that represents the current field object
 */
interface IFieldPOST {

    /** Title of the field */
    title: string;

    /** ID of the Field  */
    id?: string;

    /** Field Type enum */
    fieldType?: FieldType;

    /** Group in which field will reside */
    group: string;

    /** Description of the field */
    description: string;

    /** Is field mandatory? */
    isRequired: boolean;

    /** Default value of the field is any */
    defaultValue?: string;

    /** If the field exists */
    exists?: boolean;

    /** Is field of look up type? */
    isLookup: boolean;

    /** Lookup list Id (Only if the current field is lookup) */
    lookupListID?: string;

    /** lookup Field Title  (Only if the current field is lookup) */
    lookupColName?: string;

    /** Choices if the field is of type choice */
    choices?: string[];

    /** Schema XML of the Field (Needed to add the site column to the list) */
    schemaXml?: string;

    /** List name to which field needs to be added */
    listName?: string;

    /** Is current field needs to be added to the view */
    addToView?: boolean;

    /** View name to which current field needs to be added */
    viewName?: string;

    /** Does current field allows multi values? */
    allowMultiValues?: boolean;
}

interface IFieldGET {
    title: string;
    id: string;
    schemaXml: string;
    ok:boolean;
    exists:boolean;
    status:number;
    statusText:string;
    errorMethod:string;
}

/** Enum to represent the Field Types  */
enum FieldType {
    Text = 2,
    Note = 3,
    DateTime = 4,
    Choice = 6,
    Lookup = 7,
    Boolean = 8,
    Number = 9,
    MultiChoice = 15,
    User = 20
}

/** Interface to represent the scope of the field in SharePoint */
enum FieldScope {
    ListColumn,
    SiteColumn
}


export { FieldScope };
export { FieldType };
export { IFieldGET };
export { IFieldPOST };