/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 28 feb 2020
 * Description: This class will contain only the miscrelaneous methods required for the SPOperations
 */


import { SPBase } from './SPBase';
import { SPHttpClient } from '@microsoft/sp-http';
import { IDocResponse } from './Props/ISPCommonProps';
import { ISPPostRequest, ISPBaseResponse } from './Props/ISPBaseProps';
import { Log } from '@microsoft/sp-core-library';

const CLASS_NAME: string = `SPCommonOperations`;
/**
 * This class will contain only the miscrelaneous methods required for the SPOperations
 */
class SPCommonOperations extends SPBase {


    constructor(spHttpClient: SPHttpClient, webUrl: string, logSource: string) {
        super(spHttpClient, webUrl, logSource);
    }

    /** Use this method to get the SPCommonOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string, logSource: string): SPCommonOperations {

        return new SPCommonOperations(spHttpClient, webUrl, logSource);
    }

    /**
    * Method returns the SharePoint default icons for each file type
    * @param fileNames : array of file names for which icons are required
    */
    public async getDocIconByFiles(fileNames: string[]): Promise<IDocResponse[]> {

        let extensions: string[] = [];
        let docDetails: IDocResponse[] = [];

        try {
            // Get all the unique file extensions to redudce the number of calls
            fileNames.forEach(i => {
                let ext: string = i.slice(i.lastIndexOf('.'));
                if (extensions.indexOf(ext) == -1) {
                    extensions.push(ext);
                }
            });

            let requests: Promise<ISPBaseResponse>[] = [];

            extensions.forEach(i => {
                let url = `${this.WebUrl}/_api/web/maptoicon(filename='abc${i}', progid='', size='3')`;
                requests.push(this.spQueryGET(url));
            });

            let responses: ISPBaseResponse[] = await Promise.all(requests);

            fileNames.forEach(file => {
                // check the extension of the file
                let ext: string = file.slice(file.lastIndexOf('.'));
                // get the index of the same in the extension array
                let index: number = extensions.indexOf(ext);
                // response object for the current file
                let response: ISPBaseResponse = responses[index];
                // create the response object
                docDetails.push({ error: response.error, fileName: file, imageUrl: response.error ? undefined : `/_layouts/15/images/${response.result.value}`, ok: response.ok });
            });
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getDocIconByFiles`));
            Log.error(this.LogSource, error);
            docDetails = [{ error: error, imageUrl: undefined, fileName: undefined, ok: false }];
        }
        finally {
            return Promise.resolve(docDetails);
        }
    }

    /**
    * Method gives access to query GET 
    * @param url : query URL
    */
    public async queryGETResquest(url: string): Promise<ISPBaseResponse> {
        return await this.spQueryGET(url);
    }

    /**
     * Method gives access to query POST
     * @param postProps : ISPPostRequest object
     */
    public async queryPOSTRequest(postProps: ISPPostRequest): Promise<ISPBaseResponse> {
        return await this.spQueryPOST(postProps);
    }

    /**
     * Method gives access to query MERGE
     * @param postProps : ISPPostRequest object
     */
    public async queryMERGERequest(postProps: ISPPostRequest): Promise<ISPBaseResponse> {
        return await this.spQueryMERGE(postProps);
    }

    /**
     * Method gives access to query PATCH
     * @param postProps : ISPPostRequest object
     */
    public async queryPATCHRequest(postProps: ISPPostRequest): Promise<ISPBaseResponse> {
        return await this.spQueryPATCH(postProps);
    }
}

export { SPCommonOperations };