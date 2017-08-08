/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 1 May 2017
 * Description: This class will contain only the methods that can be used in SPBatchOperations
 */

import { SPHttpClient, SPHttpClientBatch, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISPHttpClientBatchOptions, ISPHttpClientBatchCreationOptions } from '@microsoft/sp-http';
import { ISPBaseResponse, ISPPostRequest } from './Props/ISPBaseProps';

/**
 * This class will contain only the miscrelaneous methods required for the SPOperations
 */
class SPBatchOperations {

    private spBatchCreationOptions: ISPHttpClientBatchCreationOptions;
    private spBatch: SPHttpClientBatch;

    constructor(spHttpClient: SPHttpClient, webUrl: string) {
        this.spBatchCreationOptions = { webUrl: webUrl };
        this.spBatch = spHttpClient.beginBatch(this.spBatchCreationOptions);
    }

    /**
     * Returns the SPHttpClientBatch object to execute Batch queries
     */
    public get oSPBatch(): SPHttpClientBatch {
        return this.spBatch;
    }

    /**
     * Returns the SPHttpClientResponse object with required values
     * @param url : url to be executed (it needs to be the absolute url)
     */
    public getBatchGETRequest(url: string): Promise<SPHttpClientResponse> {
        return this.oSPBatch.get(url, SPHttpClientBatch.configurations.v1);
    }

    /**
     * Returns the SPHttpClientResponse with required values
     * @param post : ISPPostRequest object. Needs to be the absolute URL.
     */
    public getBatchPOSTRequest(post: ISPPostRequest): Promise<SPHttpClientResponse> {

        const batchOps: ISPHttpClientBatchOptions = {
            body: `{ Title: 'List created with SPFx batching at', BaseTemplate: 100 }`
        };

        return this.oSPBatch.post(post.url, SPHttpClientBatch.configurations.v1, batchOps);
    }

    /**
     *  Converts SPHttpClientResponse object to ISPBaseResponse
     * @param response : Response recieved from Batch Query
     */
    public SPHttpClientResponseToSPBaseResponse(response: SPHttpClientResponse): Promise<ISPBaseResponse> {

        return response.json().then((item) => {
            let custmResponse: ISPBaseResponse = {
                ok: response.ok,
                result: item,
                status: response.status,
                statusText: response.statusText,
                errorMethod: 'SPCommonOperations.SPHttpClientResponseToSPBaseResponse',
                responseJSON: JSON.stringify(item)
            };

            return custmResponse;
        });
    }
}

export { SPBatchOperations };