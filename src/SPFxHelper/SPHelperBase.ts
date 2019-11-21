/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 16 November 2019
 * Description: This class will contain only the core methods required for the SPOperations
 */

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ISPBaseResponse, ISPPostRequest } from './Props/ISPBaseProps';
import { Log } from '@microsoft/sp-core-library';


/**
 * This file contains the base methods required to make any SharePoint operations.
 * All core interaction methods needs to be placed in this file and then to be used.
 * This class implements singleton, so only single instance is created, remember when introducing any new method.
 * Prevent using directly these methods in the webparts.
 */
class SPHelperBase {

    private spHttpClient: SPHttpClient;
    private webURL: string;
    protected errorStatus: number = -1;
    protected LOG_SOURCE: string = `SPFxHelper`;

    protected constructor(spHttpClient: SPHttpClient, webUrl: string, logSource?: string) {
        this.spHttpClient = spHttpClient;
        this.webURL = webUrl;
        this.LOG_SOURCE = logSource ? logSource : this.LOG_SOURCE;
    }

    /** return the web url */
    public get WebUrl(): string {
        return this.webURL;
    }

    /** 
     * Call this method to execute GET query.
     * Returns false if query fails else returns the response
    */
    protected async spQueryGET(query: string): Promise<ISPBaseResponse> {

        let retValue: ISPBaseResponse = undefined;
        try {

            let response: SPHttpClientResponse = await this.spHttpClient.get(query, SPHttpClient.configurations.v1);
            let responseParsed: ISPBaseResponse = await response.json();
            retValue = {
                ok: response.ok,
                result: responseParsed,
                status: response.status,
                statusText: response.statusText,
                responseJSON: response.status === 200 ? `` : JSON.stringify(responseParsed)
            }
        }
        catch (error) {
            Log.error(this.LOG_SOURCE, new Error(`Error occured in SPHelperBase.spQueryGET() method`));
            Log.error(this.LOG_SOURCE, error);
        }
        finally {
            return Promise.resolve(retValue);
        }
    }

    /** 
     * Call this method to execute POST query.
     * Returns false if query fails else returns the response
    */
    protected async spQueryPOST(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        let retValue: ISPBaseResponse = undefined;

        try {
            let options: ISPHttpClientOptions = {
                headers: { 'odata-version': '3.0' },
                body: postProps.body
            };

            retValue = await this.executePOSTRequest(postProps.url, options);
        }
        catch (error) {
            Log.error(this.LOG_SOURCE, new Error(`Error occured in SPHelperBase.spQueryPOST() method`));
            Log.error(this.LOG_SOURCE, error);
        }
        finally {
            return Promise.resolve(retValue);
        }
    }

    /** 
     * Call this method to execute MERGE query.
     * Returns false if query fails else returns the response
    */
    protected async spQueryMERGE(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        let retValue: ISPBaseResponse = undefined;

        try {
            let options: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE'
                },
                body: postProps.body
            };

            retValue = await this.executePOSTRequest(postProps.url, options);
        }
        catch (error) {
            Log.error(this.LOG_SOURCE, new Error(`Error occured in SPHelperBase.spQueryMERGE() method`));
            Log.error(this.LOG_SOURCE, error);
        }
        finally {
            return Promise.resolve(retValue);
        }
    }

    /** 
     * Call this method to execute PATCH query.
     * Returns false if query fails else returns the response
    */
    protected async spQueryPATCH(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        let retValue: ISPBaseResponse = undefined;

        try {
            let options: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'PATCH'
                },
                body: postProps.body
            };

            retValue = await this.executePOSTRequest(postProps.url, options);
        }
        catch (error) {
            Log.error(this.LOG_SOURCE, new Error(`Error occured in SPHelperBase.spQueryPATCH() method`));
            Log.error(this.LOG_SOURCE, error);
        }
        finally {
            return Promise.resolve(retValue);
        }
    }

    /** 
     * Executes the POST request for the POST/MERGE/PATCH
     */
    private async executePOSTRequest(url: string, options: ISPHttpClientOptions): Promise<ISPBaseResponse> {

        let retValue: ISPBaseResponse = undefined;

        try {

            let response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
            let responseParsed: ISPBaseResponse = await response.json();
            retValue = {
                ok: response.ok,
                result: responseParsed,
                status: response.status,
                statusText: response.statusText,
                responseJSON: response.status === 200 ? `` : JSON.stringify(responseParsed)
            }
        }
        catch (error) {
            Log.error(this.LOG_SOURCE, new Error(`Error occured in SPHelperBase.executePOSTRequest() method`));
            Log.error(this.LOG_SOURCE, error);
        }
        finally {
            return Promise.resolve(retValue);
        }
    }
}

export { SPHelperBase };
