import { IODataList } from '@microsoft/sp-odata-types';

/**
 * Interface to get the details of the list
 */
interface IListGET {
    /** If query has successfully executed */
    ok: boolean;

    /** If the list exists or not */
    exists: boolean;

    /** Details of the list if exists */
    details?: IODataList;

    /** status if any error occured (check if ok is false) */
    status: number;

    /** text of the status recieved */
    statusText: string;

    /** Method name to log where error occured */
    errorMethod: string;
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

interface IListItemResponse{
    ok: boolean;
    status: number;
    statusText: string;
    result: any[];
    errorMethod: string;
    responseJSON:string;
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