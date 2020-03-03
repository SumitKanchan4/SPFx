/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 28 Feb 2020
 * Description: This class will contain only the core methods required for the SPOperations
 */

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ISPPostRequest, ISPBaseResponse } from './Props/ISPBaseProps';
import { SPCore } from './SPCore';
import { Log } from '@microsoft/sp-core-library';

const CLASS_NAME: string = 'SPBase';
/**
 * This file contains the base methods required to make any SharePoint operations.
 * All core interaction methods needs to be placed in this file and then to be used.
 * This class implements singleton, so only single instance is created, remember when introducing any new method.
 * Prevent using directly these methods in the webparts.
 */
class SPBase {

    private spHttpClient: SPHttpClient;
    private webURL: string;
    private logSource: string = `SPFXHelper`;

    protected constructor(spHttpClient: SPHttpClient, webUrl: string, logSource: string) {
        this.spHttpClient = spHttpClient;
        this.webURL = webUrl;
        this.logSource = logSource;
    }

    /** return the web url */
    public get WebUrl(): string {
        return this.webURL;
    }

    /** return the web url */
    public get LogSource(): string {
        return this.logSource;
    }

    /** 
     * Call this method to execute GET query.
     * Returns ISPBaseResponse object with the response of the query
    */
    protected async spQueryGET(query: string): Promise<ISPBaseResponse> {

        let result: ISPBaseResponse;
        let respObj: any;

        try {
            // Create header
            let options: ISPHttpClientOptions = {
                headers: {
                    'Accept': 'application/json',
                    'odata-version': ''
                }
            };
            // get the response of the query
            let response: SPHttpClientResponse = await this.spHttpClient.get(query, SPHttpClient.configurations.v1, options);
            respObj = await response.json();

            if (response.ok) {
                result = { error: undefined, ok: true, result: respObj };
            }
            else {
                result = { error: new Error(`${respObj["odata.error"]["message"]["value"]}`), ok: false, result: undefined };
                Log.error(this.logSource, new Error(`Error occured in SPHelperBase.spQueryGET`));
                Log.error(this.LogSource, result.error);
            }

        }
        catch (error) {
            Log.error(this.logSource, new Error(`Error occured in SPHelperBase.spQueryGET`));
            Log.error(this.logSource, new Error(`Query: ${query}`));
            Log.error(this.logSource, error);
            result = { error: error, ok: false, result: undefined };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /** 
     * Call this method to execute POST query.
     * Returns the response on the execution of the query
    */
    protected async spQueryPOST(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        let result: ISPBaseResponse;
        try {
            let options: ISPHttpClientOptions = { headers: { 'odata-version': '3.0' } };

            if (!SPCore.isEmptyString(postProps.body)) {
                options["body"] = postProps.body;
            }

            result = await this.executePOSTRequest(postProps.url, options);
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.spQueryPOST`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error, result: undefined };
        }
        finally {
            return result;
        }
    }

    /** 
    * Call this method to execute MERGE query.
    * Returns ISPBaseResponse object with the response of the query
   */
    protected async spQueryMERGE(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        let result: ISPBaseResponse;
        try {
            let options: ISPHttpClientOptions = {};

            if (!SPCore.isEmptyString(postProps.body)) {
                options = {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': '',
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE'
                    },
                    body: postProps.body
                };
            }
            else {
                options = {
                    headers: { 'odata-version': '3.0' },
                };
            }

            result = await this.executePOSTRequest(postProps.url, options);
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.spQueryMERGE`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error, result: undefined };
        }
        finally {
            return result;
        }
    }

    /** 
     * Call this method to execute PATCH query.
     * Returns ISPBaseResponse object with the response of the query
    */
    protected async spQueryPATCH(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        let result: ISPBaseResponse;

        try {
            let options: ISPHttpClientOptions = {};

            if (!SPCore.isEmptyString(postProps.body)) {
                options = {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=verbose',
                        'odata-version': '',
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'PATCH'
                    },
                    body: postProps.body
                };
            }
            else {
                options = {
                    headers: { 'odata-version': '3.0' },
                };
            }

            result = await this.executePOSTRequest(postProps.url, options);
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.spQueryPATCH`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error, result: undefined };
        }
        finally {
            return result;
        }
    }

    /** 
     * Executes the request for the POST/MERGE/PATCH
     */
    private async executePOSTRequest(url: string, options: ISPHttpClientOptions): Promise<ISPBaseResponse> {

        let result: ISPBaseResponse;

        try {

            let response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
            let respObj = await response.json();

            if (response.ok) {
                result = { ok: true, result: respObj, error: undefined };
            }
            else {
                result = { ok: false, error: new Error(`${respObj["odata.error"]["message"]["value"]}`), result: undefined };
                Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.executePOSTRequest`));
                Log.error(this.LogSource, result.error);
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.executePOSTRequest`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error, result: undefined };
        }
        finally {
            return result;
        }
    }
}

export { SPBase };
