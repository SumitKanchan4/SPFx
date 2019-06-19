/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 1 May 2017
 * Description: This class will contain only the core methods required for the SPOperations
 */

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ISPBaseResponse, ISPPostRequest } from './Props/ISPBaseProps';
import { SPHelperCommon } from './SPHelperCommon';
import { SPLogger } from './SPLogger';


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

    protected constructor(spHttpClient: SPHttpClient, webUrl: string) {
        this.spHttpClient = spHttpClient;
        this.webURL = webUrl;
    }

    /** return the web url */
    public get WebUrl(): string {
        return this.webURL;
    }

    /** 
     * Call this method to execute GET query.
     * Returns false if query fails else returns the response
    */
    protected spQueryGET(query: string): Promise<ISPBaseResponse> {
        try {
            return this.spHttpClient.get(query,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'odata-version': ''
                    }
                }).then((response: SPHttpClientResponse) => {

                    return response.json().then((item) => {
                        var custmResponse: ISPBaseResponse = {
                            ok: response.ok,
                            result: item,
                            status: response.status,
                            statusText: response.statusText,
                            errorMethod: 'SPHelperBase.spQueryGET',
                            responseJSON: JSON.stringify(item)
                        };

                        return custmResponse;
                    });
                });
        }
        catch (error) {
            SPLogger.logError(error as Error);
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPHelperBase.spQueryGET'
            });
        }
    }

    /** 
     * Call this method to execute POST query.
     * Returns false if query fails else returns the response
    */
    protected spQueryPOST(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        try {
            var options: ISPHttpClientOptions = {};

            if (!SPHelperCommon.isStringNullOrEmpty(postProps.body)) {
                options = {
                    headers: { 'odata-version': '3.0' },
                    body: postProps.body
                };
            }
            else {
                options = {
                    headers: { 'odata-version': '3.0' },
                };
            }
            return this.executePOSTRequest(postProps.url, options).then((postResponse) => {
                return postResponse;
            });
        }
        catch (error) {
            SPLogger.logError(error as Error);
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPHelperBase.spQueryPOST',
                responseJSON: JSON.stringify(error)
            });
        }
    }

    /** 
     * Call this method to execute MERGE query.
     * Returns false if query fails else returns the response
    */
    protected spQueryMERGE(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        try {
            var options: ISPHttpClientOptions = {};

            if (!SPHelperCommon.isStringNullOrEmpty(postProps.body)) {
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

            return this.executePOSTRequest(postProps.url, options).then((mergeResponse) => {
                return mergeResponse;
            });
        }
        catch (error) {
            SPLogger.logError(error as Error);
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPHelperBase.spQueryMERGE'
            });
        }
    }

    /** 
     * Call this method to execute PATCH query.
     * Returns false if query fails else returns the response
    */
    protected spQueryPATCH(postProps: ISPPostRequest): Promise<ISPBaseResponse> {

        try {
            var options: ISPHttpClientOptions = {};

            if (!SPHelperCommon.isStringNullOrEmpty(postProps.body)) {
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

            return this.executePOSTRequest(postProps.url, options).then((patchResponse) => {
                return patchResponse;
            });
        }
        catch (error) {
            SPLogger.logError(error as Error);
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPHelperBase.spQueryPATCH',
                responseJSON: JSON.stringify(error)
            });
        }
    }

    /** 
     * Executes the POST request for the POST/MERGE/PATCH
     */
    private executePOSTRequest(url: string, options: ISPHttpClientOptions): Promise<ISPBaseResponse> {
        try {

            return this.spHttpClient.post(url, SPHttpClient.configurations.v1, options)
                .then((response: SPHttpClientResponse) => {

                    return response.json().then((item) => {
                        var custmResponse: ISPBaseResponse = {
                            ok: response.ok,
                            result: item,
                            status: response.status,
                            statusText: response.statusText,
                            errorMethod: 'SPHelperBase.executePOSTRequest',
                            responseJSON: JSON.stringify(item)
                        };

                        return custmResponse;
                    });
                });
        }
        catch (error) {
            SPLogger.logError(error as Error);
            return Promise.resolve({
                ok: false,
                result: error,
                status: this.errorStatus,
                statusText: error.message,
                errorMethod: 'SPHelperBase.executePOSTRequest',
                responseJSON: JSON.stringify(error)
            });
        }
    }
}

export { SPHelperBase };
