/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 1 May 2017
 * Description: This class will contain only the miscrelaneous methods required for the SPOperations
 */


import { SPHelperBase } from './SPHelperBase';
import { SPHelperCommon } from './SPHelperCommon';
import { SPHttpClient } from '@microsoft/sp-http';
import { IDoc, IDocResponse } from './Props/ISPCommonProps';
import { ISPBaseResponse, ISPPostRequest } from './Props/ISPBaseProps';
import { SPLogger } from './SPLogger';
import { Log } from '@microsoft/sp-core-library';


/**
 * This class will contain only the miscelaneous methods required for the SPOperations
 */
class SPCommonOperations extends SPHelperBase {

    private docDetails: IDocResponse[] = [];

    constructor(spHttpClient: SPHttpClient, webUrl: string, logSource?: string) {
        super(spHttpClient, webUrl, logSource);
    }

    /** Use this method to get the SPCommonOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string, logSource?: string): SPCommonOperations {

        return new SPCommonOperations(spHttpClient, webUrl, logSource);
    }

    /**
     * Method returns the SharePoint default icons for each file type
     * @param fileNames : array of files for which icons are required
     */
    public async getDocIconByFiles(files: IDoc[]): Promise<IDocResponse[]> {
        this.docDetails = [];
        try {

            if (files && files.length > 0) {
                this.docDetails = await this.getIcons(files);
            }
        }
        catch (error) {
            Log.error(this.LOG_SOURCE, new Error(`Error occured in SPCommonOperations.getDocIconByFiles() method`));
            Log.error(this.LOG_SOURCE, error);
        }
        finally {
            return Promise.resolve(this.docDetails);
        }
    }

    /**
     * Private method to get the document icons.
     * Method is recursive in nature so made this private
     * @param fileNames : Filenames of which icon is needed
     */
    private async getIcons(files: IDoc[]): Promise<IDocResponse[]> {

        let fileName: string = undefined;
        let file: IDoc = files.pop();

        if (file.fileName) {
            fileName = file.fileName;
        }
        // Get the fileName from file url
        else if (!SPHelperCommon.isStringNullOrEmpty(file.fileUrl)) {
            var fileNameIndex = file.fileUrl.lastIndexOf("/") + 1;
            fileName = file.fileUrl.substr(fileNameIndex);
        }

        // If filename is empty
        if (SPHelperCommon.isStringNullOrEmpty(fileName)) {
            this.docDetails.push(
                {
                    ok: false,
                    status: this.errorStatus,
                    statusText: 'Both file name or file url cannot be empty',
                    image: null,
                    fileName: fileName,
                    fileUrl: file.fileUrl,
                    id: file.id,
                    success: false,
                    errorMethod: 'SPCommonOperations.getIcons'
                });
            if (files.length > 0) {
                return this.getIcons(files);
            } else {
                return Promise.resolve(this.docDetails);
            }
        }
        else {
            var url = `${this.WebUrl}/_api/web/maptoicon(filename='${fileName}', progid='', size='3')`;
            try {

                return this.spQueryGET(url).then((response) => {

                    if (response.ok && !SPHelperCommon.isNull(response.result)) {

                        this.docDetails.push({
                            fileName: fileName,
                            fileUrl: file.fileUrl,
                            id: file.id,
                            success: true,
                            image: `/_layouts/15/images/${response.result.value}`,
                            errorMethod: 'SPCommonOperations.getIcons',
                            ok: response.ok,
                            status: response.status,
                            statusText: response.statusText
                        });
                    }
                    else {
                        this.docDetails.push(
                            {
                                ok: response.ok,
                                status: response.status,
                                statusText: response.ok ? response.statusText : 'Could not recieve the result.',
                                image: null,
                                fileName: fileName,
                                fileUrl: file.fileUrl,
                                id: file.id,
                                success: false,
                                errorMethod: 'SPCommonOperations.getIcons'
                            });
                    }

                    if (files.length > 0) {
                        return this.getIcons(files);
                    } else {
                        return Promise.resolve(this.docDetails);
                    }
                });
            } catch (error) {
                SPLogger.logError(error as Error);
                this.docDetails.push(
                    {
                        ok: false,
                        status: this.errorStatus,
                        statusText: error.message,
                        image: null,
                        fileName: file.fileName,
                        fileUrl: file.fileUrl,
                        id: file.id,
                        success: false,
                        errorMethod: 'SPCommonOperations.getIcons'
                    });
                return Promise.resolve(this.docDetails);
            }
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