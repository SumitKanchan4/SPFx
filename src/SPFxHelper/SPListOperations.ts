import { SPBase } from './SPBase';
import { IListGET, IListPOST, IListItemResponse, IItem } from './Props/ISPListProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { ISPPostRequest, ISPBaseResponse } from './Props/ISPBaseProps';
import { IListItemsResponse, BaseTemplate, ILibraryItemResponse, ILibraryItemsResponse } from './Props/ISPListProps';
import { Log } from '@microsoft/sp-core-library';
import { SPCore } from './SPCore';

const CLASS_NAME: string = `SPListOperations`;
// const CLASS_NAME: string = `SPListOperations`;
/**
 * This class will contain only the List-Library specific methods required for the SPOperations
 */
class SPListOperations extends SPBase {

    constructor(spHttpClient: SPHttpClient, webUrl: string, logSource: string) {
        super(spHttpClient, webUrl, logSource);
    }

    /** Use this method to get the SPListOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string, logSource: string): SPListOperations {

        return new SPListOperations(spHttpClient, webUrl, logSource);
    }

    /**
     * Method returns the list details if exists. 
     * Check property exists to check if list exists or not.
     * This will check the list existence only in current web.
     * @param lstTitle : Provide the title of the list
     */
    public async getListByTitle(lstTitle: string): Promise<IListGET> {

        // if list title is blank, return with error
        if (SPCore.isEmptyString(lstTitle)) { return Promise.resolve({ exists: false, ok: false, error: new Error(`List Title cannot be blank`) }); }

        let result: IListGET;

        try {
            // Create query
            let query: string = `${this.WebUrl}/_api/web/lists/getByTitle('${lstTitle}')`;
            // get the response based
            let response: ISPBaseResponse = await this.spQueryGET(query);

            if (response.ok) {
                result = { exists: !!response.result, ok: true, details: response.result };
            }
            else {
                // Set ok,error if the 404 is found in error
                result = { exists: false, ok: false, error: response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getListByTitle`));
            Log.error(this.LogSource, error);
            result = { exists: false, ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Checks for the list existence by title and Base Template
     *  else creates the list based on the base template and the properties defined.
     * @param lstDetail : Details of the list 
     */
    public async createList(lstDetail: IListPOST): Promise<IListGET> {

        // Return the response if the object is null
        if (!lstDetail) { return Promise.resolve({ exists: false, ok: false, error: new Error(`List detail object cannot be null`) }); }

        let result: IListGET;

        try {
            // Check if the list with the same title exists
            Log.verbose(this.LogSource, `Check if the list with the same title exists`);
            let lst: IListGET = await this.getListByTitle(lstDetail.title);
            if (lst.exists && lst.details["BaseTemplate"] === lstDetail.baseTemplate) {
                Log.verbose(this.LogSource, `List exists with same title and base template, so not created new`);
                result = lst;
            }
            else {
                let createdList = await this.spQueryPOST(this.getListMetadata(lstDetail));
                if (createdList.ok) {
                    result = { ok: true, exists: true, details: createdList["result"] };
                }
                else {
                    result = { exists: false, ok: false, error: createdList.error };
                }
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.createList`));
            Log.error(this.LogSource, error);
            result = { exists: false, ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method returns the items from the respective list 
     * @param listTitle : Title of the list from where  data is required
     * @param query : Query to filter the results. Include '?' in the start.
     */
    public async getListItemsByQuery(listTitle: string, query: string): Promise<IListItemsResponse> {

        return await this.getListItemsBase(listTitle, query);
    }

    /**
     * Method returns the items from the respective list 
     * @param listTitle : Title of the list from where  data is required
     * @param query : Query srtarting from '?' if any
     */
    private async getListItemsBase(listTitle: string, query?: string): Promise<IListItemsResponse> {

        let result: IListItemsResponse;
        try {
            let url = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/Items`;
            url = SPCore.isEmptyString(query) ? url : `${url}${query}`;

            let response: ISPBaseResponse = await this.spQueryGET(url);
            if (response.ok) {
                result = { ok: true, result: response.result.value, nextLink: !!response.result["odata.nextLink"] ? response.result["odata.nextLink"] : undefined };
            }
            else {
                result = { ok: false, error: response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getListItemsBase`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method returns the item of the respective list by Item ID
     * @param listTitle : Title of the list from where  data is required
     * @param itemID : ID of the item
     */
    public async getListItemByID(listTitle: string, itemID: string): Promise<IListItemResponse> {

        let result: IListItemResponse = { ok: false };
        try {
            
            let url = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/Items(${itemID})`;

            let response: ISPBaseResponse = await this.spQueryGET(url);
            if (response.ok) {
                result = { ok: true, result: response.result };
            }
            else {
                result = { ok: false, error: response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getListItemByID`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }

    }

    /**
     * Method returns all the items based on max rows
     * @param listTitle : List title
     * @param rowCount : number of rows to return. Define 0 to get default max rows
     */
    public async getListItems(listTitle: string, rowCount: number): Promise<IListItemsResponse> {

        let query = rowCount && rowCount > 0 ? `?$top=${rowCount}` : undefined;
        return await this.getListItemsBase(listTitle, query);
    }

    /**
     * Method returns the items from the respective list 
     * @param libraryTitle : Title of the library from where  data is required
     * @param query : Query srtarting from '?' if any
     */
    private async getLibraryItemsBase(libraryTitle: string, query?: string): Promise<ILibraryItemsResponse> {

        let result: IListItemsResponse;
        try {
            let url = `${this.WebUrl}/_api/web/lists/getByTitle('${libraryTitle}')/Files`;
            url = SPCore.isEmptyString(query) ? url : `${url}${query}`;

            let response: ISPBaseResponse = await this.spQueryGET(url);
            if (response.ok) {
                result = { ok: true, result: response.result.value, nextLink: !!response.result["odata.nextLink"] ? response.result["odata.nextLink"] : undefined };
            }
            else {
                result = { ok: false, error: response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getLibraryItemsBase`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method returns the item of the respective list by Item ID
     * @param libraryTitle : Title of the library from where  data is required
     * @param fileId : ID of the file (do not confuse with the item)
     */
    public async getLibraryItemByFileID(libraryTitle: string, fileId: string): Promise<ILibraryItemResponse> {

        let result: ILibraryItemResponse = { ok: false };
        try {
           
            let url = `${this.WebUrl}/_api/web/lists/getByTitle('${libraryTitle}')/Files('${fileId}')`;

            let response: ISPBaseResponse = await this.spQueryGET(url);
            if (response.ok) {
                result = { ok: true, result: response.result };
            }
            else {
                result = { ok: false, error: response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getLibraryItemByFileID`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method returns all the items based on max rows
     * @param libraryTitle : library title
     * @param rowCount : number of rows to return. Define 0 to get default max rows
     */
    public async getLibraryItems(libraryTitle: string, rowCount: number): Promise<ILibraryItemsResponse> {

        let query = rowCount && rowCount > 0 ? `?$top=${rowCount}` : undefined;
        return await this.getLibraryItemsBase(libraryTitle, query);
    }

    /**
     * Method creates the item in the list based on the body structure to the respective list
     * @param listTitle : Title of the list where item needs to be created
     * @param body : Body of the item 
     */
    public async createListItem(listTitle: string, itemValues: IItem[]): Promise<IListItemResponse> {

        let result: IListItemResponse = { ok: false };
        try {

            let listProps: IListGET = await this.getListByTitle(listTitle);

            if (listProps.exists) {
                let body = {
                    __metadata: {
                        'type': `${listProps.details.ListItemEntityTypeFullName}`
                    }
                };

                itemValues.forEach(i => {
                    body[`${i.fieldName}`] = i.fieldValue;
                });

                let response: ISPBaseResponse = await this.spQueryPOST({ body: JSON.stringify(body), url: `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/items` });
                result = { ok: response.ok, error: response.error, result: response.result };
            }
            else {
                result.error = listProps.error;
                result.ok = listProps.ok;
            }

        } catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.createListItem`));
            Log.error(this.LogSource, error);
            result.error = error;
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Creates the folder in the document library
     * @param docLib : Title of the document Library (Internal name and not the title) where folder needs to be created
     * @param folderName : Folder name
     */
    public async createFolderInDocLib(docLib: string, folderName: string): Promise<IListItemResponse> {

        let result: IListItemResponse = { ok: false };
        try {
            let response: ISPBaseResponse = await this.spQueryPOST({ body: undefined, url: `${this.WebUrl}/_api/web/folders/add('${docLib}/${folderName}')` });
            result = { error: response.error, ok: response.ok, result: response.result };
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.createFolderInDocLib`));
            Log.error(this.LogSource, error);
            result.error = error;
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method creates the folder in the list 
     * @param listTitle : Title fo the list where folder needs to be created
     * @param folderName : Name of the folder needs to be created
     */
    public async createFolderInList(listTitle: string, folderName: string): Promise<IListItemResponse> {

        let result: IListItemResponse;
        try {

            let item: IItem[] = [
                { fieldName: 'Title', fieldValue: `${folderName}` },
                { fieldName: 'FileLeafRef', fieldValue: `${folderName}` },
                { fieldName: 'FileSystemObjectType', fieldValue: '1' },
                { fieldName: 'ContentTypeId', fieldValue: '0x0120' },
            ];

            result = await this.createListItem(listTitle, item);

            if (result.ok) {
                let body: string = JSON.stringify({
                    '__metadata': {
                        'type': `${result.result["odata.type"]}`
                    },
                    'Title': `${folderName}`,
                    'FileLeafRef': `${folderName}`
                });

                let urlPOST = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${result.result["Id"]})`;
                this.spQueryPATCH({ body: body, url: urlPOST });
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.createFolderInList`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Updates the item in the list in the current web
     * @param listTitle : Title of the list where item needs to be updated
     * @param itemId : ID of the item which needs to be updated
     * @param body : Body template of the item which needs to be updated
     */
    public async updateListItem(listTitle: string, itemId: string, itemValues: IItem[]): Promise<IListItemResponse> {

        let result: IListItemResponse = { ok: false };

        try {

            let listProps: IListGET = await this.getListByTitle(listTitle);

            if (listProps.exists) {

                let body = {
                    __metadata: {
                        'type': `${listProps.details.ListItemEntityTypeFullName}`
                    }
                };

                itemValues.forEach(i => {
                    body[`${i.fieldName}`] = i.fieldValue;
                });

                let response: ISPBaseResponse = await this.spQueryMERGE({ body: JSON.stringify(body), url: `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})` });
                result = { ok: response.ok, error: response.error, result: response.result };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.updateListItem`));
            Log.error(this.LogSource, error);
            result.error = error;
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Returns the metadata for the list creation
     * @param lstDetail 
     * @param webURL 
     */
    public getListMetadata(lstDetail: IListPOST): ISPPostRequest {

        let postListData: string = `{ '__metadata': { 'type': 'SP.List' }, 'BaseTemplate': ${lstDetail.baseTemplate},'EnableFolderCreation': ${lstDetail.allowFolder}, 'ContentTypesEnabled': ${lstDetail.allowContentTypes}, 'AllowContentTypes': ${lstDetail.allowContentTypes}, 'Description': '${lstDetail.description}', 'Title':'${lstDetail.title}'}`;
        let postURL: string = `${this.WebUrl}/_api/web/lists`;

        let post: ISPPostRequest = {
            body: postListData,
            url: postURL
        };

        return post;
    }

    /**
     * 
     * @param baseTemplate Returns the list details based on the Base Template
     */
    public async getListsDetailsByBaseTemplateID(baseTemplate: BaseTemplate): Promise<ISPBaseResponse> {

        let url = `${this.WebUrl}/_api/web/lists?$filter=BaseTemplate eq ${baseTemplate}`;

        return await this.spQueryGET(url);
    }

    /**
     * Returns the views associated with the list.
     * @param listTitle :Title of the list
     */
    public async getViewsByList(listTitle: string): Promise<ISPBaseResponse> {

        let url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/views/`;

        return await this.spQueryGET(url);
    }

    /**
     * Returns all the content types associated with the list
     * @param listTitle : Title of the list
     */
    public async getContentTypesByList(listTitle: string): Promise<ISPBaseResponse> {

        var url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/contentTypes/`;
        return await this.spQueryGET(url);
    }
}

export { SPListOperations };
