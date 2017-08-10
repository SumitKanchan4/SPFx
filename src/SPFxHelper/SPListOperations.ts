import { SPHelperBase } from './SPHelperBase';
import { SPHelperCommon } from './SPHelperCommon';
import { IListGET, IListPOST } from './Props/ISPListProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { ISPBaseResponse, ISPPostRequest } from './Props/ISPBaseProps';
import { IListItemResponse, BaseTemplate } from './Props/ISPListProps';

/**
 * This class will contain only the List-Library specific methods required for the SPOperations
 */
class SPListOperations extends SPHelperBase {

    private static instance: SPListOperations;

    private constructor(spHttpClient: SPHttpClient, webUrl: string) {
        super(spHttpClient, webUrl);
    }

    /** Use this method to get the SPListOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string): SPListOperations {

        SPListOperations.instance = SPHelperCommon.isNull(SPListOperations.instance) ? new SPListOperations(spHttpClient, webUrl) : SPListOperations.instance;

        return SPListOperations.instance;

    }

    /**
     * Method returns the list details if exists. 
     * Check property exists to check if list exists or not.
     * This will check the list existence only in current web.
     * @param lstTitle : Provide the title of the list
     */
    public getListByTitle(lstTitle: string): Promise<IListGET> {
        try {
            var query: string = `${this.WebUrl}/_api/web/lists?$filter=Title eq '${lstTitle}'`;

            return this.spQueryGET(query).then((response) => {

                var listDetails: IListGET = {
                    ok: response.ok,
                    status: response.status,
                    statusText: response.statusText,
                    exists: response.result.value.length > 0 ? true : false,
                    details: response.result.value.length > 0 ? response.result.value[0] : null,
                    errorMethod: response.errorMethod
                };

                return Promise.resolve(listDetails);
            });
        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                status: this.errorStatus,
                statusText: error.message,
                exists: false,
                details: null,
                errorMethod: 'SPListOperations.getListByTitle'
            });
        }
    }

    /**
     * Checks for the list existence by title and Base Template
     *  else creates the list based on the base template and the properties defined.
     * @param lstDetail : Details of the list 
     */
    public createList(lstDetail: IListPOST): Promise<IListGET> {

        if (!SPHelperCommon.isNull(lstDetail)) {

            try {

                // Check if the list/library exists
                return this.getListByTitle(lstDetail.title).then((lstResponse) => {

                    // Check if the query executes successfully
                    if (lstResponse.ok) {

                        // Checking if list exists with the same base template and title then return the list details else create and then return the list details
                        if (lstResponse.exists && lstResponse.details.BaseTemplate == lstDetail.baseTemplate) {
                            return lstResponse;
                        }
                        else {
                            return this.spQueryPOST(this.getListMetadata(lstDetail, this.WebUrl)).then((lstCreateResp) => {

                                var listDetails: IListGET = {
                                    ok: lstCreateResp.ok,
                                    status: lstCreateResp.status,
                                    statusText: lstCreateResp.statusText,
                                    exists: lstCreateResp.result.value.length > 0 ? true : false,
                                    details: lstCreateResp.result.value.length > 0 ? lstCreateResp.result.value[0] : null,
                                    errorMethod: 'SPListOperations.createList'
                                };

                                return Promise.resolve(listDetails);
                            });
                        }

                    }
                    else {
                        // return the response as it is to track the error
                        return (lstResponse);
                    }
                });
            }
            catch (error) {
                return Promise.resolve({
                    ok: false,
                    status: this.errorStatus,
                    statusText: error.message,
                    exists: false,
                    details: null,
                    errorMethod: 'SPListOperations.createList'
                });
            }
        }
    }

    /**
     * Method returns the items from the respective list 
     * @param listTitle : Title of the list from where  data is required
     * @param query : Query to filter the results. Include '?' in the start.
     */
    public getListItemsByQuery(listTitle: string, query: string): Promise<IListItemResponse> {

        try {
            return this.getListItemsBase(listTitle, query).then((response) => {
                return response;
            });
        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.getListItemsByQuery',
                responseJSON: JSON.stringify(error)
            });
        }
    }

    /**
     * Method returns the items from the respective list 
     * @param listTitle : Title of the list from where  data is required
     * @param query : Query srtarting from '?' if any
     */
    private getListItemsBase(listTitle: string, query?: string): Promise<IListItemResponse> {

        try {
            var url = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/Items`;
            url = SPHelperCommon.isStringNullOrEmpy(query) ? url : `${url}${query}`;

            return this.spQueryGET(url).then((response) => {
                var itemsResponse: IListItemResponse = {
                    errorMethod: 'SPListOperations.getListItems',
                    ok: response.ok,
                    result: SPHelperCommon.isNull(response.result) ? null : response.result.value,
                    status: response.status,
                    statusText: response.statusText,
                    responseJSON: (SPHelperCommon.isNull(response.result) || SPHelperCommon.isNull(response.result.value)) ? JSON.stringify(response) : null
                };
                return Promise.resolve(itemsResponse);
            });
        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.getListItems',
                responseJSON: JSON.stringify(error)
            });
        }
    }

    /**
     * Method returns the item of the respective list by Item ID
     * @param listTitle : Title of the list from where  data is required
     * @param itemID : ID of the item
     */
    public getListItemByID(listTitle: string, itemID: string): Promise<IListItemResponse> {

        try {
            var query = `(${itemID})`;

            return this.getListItemsBase(listTitle, query).then((response) => {
                return response;
            });
        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.getListItemByID',
                responseJSON: JSON.stringify(error)
            });
        }
    }

    /**
     * Method returns all the items based on max rows
     * @param listTitle : List title
     * @param rowCount : number of rows to return. Define 0 to get default max rows
     */
    public getListItems(listTitle: string, rowCount: number): Promise<IListItemResponse> {

        try {
            var query = rowCount > 0 ? `?$top=${rowCount}` : null;
            return this.getListItemsBase(listTitle, query).then((response) => {
                return response;
            });
        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.getListItems',
                responseJSON: JSON.stringify(error)
            });
        }
    }

    /**
     * Method creates the item in the list based on the body structure to the respective list
     * @param listTitle : Title of the list where item needs to be created
     * @param body : Body of the item 
     */
    public createListItem(listTitle: string, body: string): Promise<ISPBaseResponse> {
        try {

            return this.spQueryPOST({
                body: body,
                url: `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/items`
            }).then((response) => {
                return response;
            });

        } catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.createListItem'
            });
        }
    }

    /**
     * Creates the folder in the document library
     * @param docLib : Title of the document Library where folder needs to be created
     * @param folderName : Folder name in the format:DocLib/FolderName
     */
    public createFolderInDocLib(docLib: string, folderName: string): Promise<ISPBaseResponse> {
        try {
            return this.spQueryPOST({
                body: null,
                url: `${this.WebUrl}/_api/web/folders/add('${docLib}/${folderName}')`
            }).then((response) => {
                return response;
            });
        } catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.createFolderInDocLib'
            });
        }
    }

    /**
     * Method creates the folder in the list 
     * @param listTitle : Title fo the list where folder needs to be created
     * @param folderName : Name of the folder needs to be created
     */
    public createFolderInList(listTitle: string, folderName: string): Promise<ISPBaseResponse> {
        try {

            return this.getListByTitle(listTitle).then((lstDetails) => {

                // Check if the list exists and query executed successfully
                if (lstDetails.ok && lstDetails.exists) {

                    var body: string = JSON.stringify({
                        "__metadata": { type: `${lstDetails.details.ListItemEntityTypeFullName}` },
                        Title: `${folderName}`,
                        FileLeafRef: `${folderName}`,
                        FileSystemObjectType: `1`,
                        ContentTypeId: "0x0120"
                    });

                    return this.createListItem(listTitle, body).then((response) => {

                        // Check if the response is ok and response contains the required data
                        if (response.ok && !SPHelperCommon.isNull(response.result.Id)) {

                            var itemId: string = response.result.Id;

                            body = JSON.stringify({
                                '__metadata': {
                                    'type': `${lstDetails.details.ListItemEntityTypeFullName}`
                                },
                                'Title': `${folderName}`,
                                'FileLeafRef': `${folderName}`
                            });

                            var urlPOST = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;

                            // IF the folders are allowed to be created it will patch the item and will show them as folder
                            this.spQueryPATCH({ body: body, url: urlPOST });

                            return Promise.resolve(response);
                        }
                        else {
                            return Promise.resolve({
                                ok: false,
                                result: { error: JSON.stringify(response.result) },
                                status: this.errorStatus,
                                statusText: `Error creating item in List: ${listTitle}`,
                                errorMethod: 'SPListOperations.createFolderInList',
                                responseJSON: JSON.stringify(response)
                            });
                        }
                    });
                }
                else {
                    return Promise.resolve({
                        ok: false,
                        result: { error: `Invalid List name: ${listTitle}` },
                        status: this.errorStatus,
                        statusText: `Invalid List name: ${listTitle}`,
                        errorMethod: 'SPListOperations.createFolderInList',
                        responseJSON: JSON.stringify(lstDetails)
                    });
                }
            });

        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.createFolderInList'
            });
        }
    }

    /**
     * Updates the item in the list in the current web
     * @param listTitle : Title of the list where item needs to be updated
     * @param itemId : ID of the item which needs to be updated
     * @param body : Body template of the item which needs to be updated
     */
    public updateListItem(listTitle: string, itemId: string, body: string): Promise<ISPBaseResponse> {
        try {
            var urlPOST = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;

            return this.spQueryMERGE({ body: body, url: urlPOST }).then((response) => {
                return response;
            });
        }
        catch (error) {
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPListOperations.updateListItem'
            });
        }
    }

    /**
     * Returns the metadata for the list creation
     * @param lstDetail 
     * @param webURL 
     */
    public getListMetadata(lstDetail: IListPOST, webURL: string): ISPPostRequest {

        var postListData: string = `{ '__metadata': { 'type': 'SP.List' }, 'BaseTemplate': ${lstDetail.baseTemplate},'EnableFolderCreation': ${lstDetail.allowFolder}, 'ContentTypesEnabled': ${lstDetail.allowContentTypes}, 'AllowContentTypes': ${lstDetail.allowContentTypes}, 'Description': '${lstDetail.description}', 'Title':'${lstDetail.title}'}`;
        var postURL: string = `${webURL}/_api/web/lists`;

        var post: ISPPostRequest = {
            body: postListData,
            url: postURL
        };

        return post;
    }

    /**
     * 
     * @param baseTemplate Returns the list details based on the Base Template
     */
    public getListsDetailsByBaseTemplateID(baseTemplate: BaseTemplate): Promise<ISPBaseResponse> {

        var url = `${this.WebUrl}/_api/web/lists?$filter=BaseTemplate eq ${baseTemplate}`;

        return this.spQueryGET(url).then((response) => {
            return response;
        });
    }
}

export { SPListOperations };