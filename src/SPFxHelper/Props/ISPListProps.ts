/**
 * Interface to get the details of the list
 */
interface IListGET {
    ok: boolean;
    exists: boolean;
    details?: any;
    error?: Error;
}

/**
 * Interface to POST the list
 */
interface IListPOST {
    title: string;
    allowContentTypes: boolean;
    allowFolder: boolean;
    id?: string;
    baseTemplate: BaseTemplate;
    description: string;
}

interface IListItemsResponse {
    ok: boolean;
    result?: any[];
    error?: Error;
    nextLink?: string;
}


interface IListItemResponse {
    ok: boolean;
    result?: any;
    error?: Error;
}

interface IItem {
    fieldName: string;
    fieldValue: string;
}

/**
 * Enum for the base templates supported right now
 */
enum BaseTemplate {
    GenericList = 100,
    DocumentLibrary = 101,
    PictureLibrary = 109
}


export { BaseTemplate };
export { IListGET };
export { IListPOST };
export { IListItemResponse };
export { IListItemsResponse };
export { IItem };