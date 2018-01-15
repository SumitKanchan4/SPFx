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


/**
 * This class will contain only the miscrelaneous methods required for the SPOperations
 */
class SPCommonOperations extends SPHelperBase {

    private static instance: SPCommonOperations;
    private docDetails: IDocResponse[] = [];

    constructor(spHttpClient: SPHttpClient, webUrl: string) {
        super(spHttpClient, webUrl);
    }

    /** Use this method to get the SPCommonOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string): SPCommonOperations {

        SPCommonOperations.instance = SPHelperCommon.isNull(SPCommonOperations.instance) ? new SPCommonOperations(spHttpClient, webUrl) : SPCommonOperations.instance;

        return SPCommonOperations.instance;

    }

    /**
     * Method returns the SharePoint default icons for each file type
     * @param fileNames : array of files for which icons are required
     */
    public getDocIconByFiles(files: IDoc[]): Promise<IDocResponse[]> {
        this.docDetails = [];
        try {

            if (!SPHelperCommon.isNull(files) && files.length > 0) {
                return this.getIcons(files).then(() => {
                    return Promise.resolve(this.docDetails);
                });
            }
            else {
                this.docDetails.push(
                    {
                        ok: false,
                        status: this.errorStatus,
                        statusText: 'files array cannot be empty',
                        image: null,
                        fileName: null,
                        fileUrl: null,
                        id: null,
                        success: false,
                        errorMethod: 'SPCommonOperations.getDocIconByFileName'
                    });
                return Promise.resolve(this.docDetails);
            }
        }
        catch (error) {
            this.docDetails.push(
                {
                    ok: false,
                    status: this.errorStatus,
                    statusText: `Recieved Props:${JSON.stringify(files)} : Error: ${error.message}`,
                    image: null,
                    fileName: null,
                    id: null,
                    success: false,
                    fileUrl: null,
                    errorMethod: 'SPCommonOperations.getDocIconByFileName'
                });
            return Promise.resolve(this.docDetails);
        }
    }

    /**
     * Private method to get the document icons.
     * Method is recursive in nature so made this private
     * @param fileNames : Filenames of which icon is needed
     */
    private getIcons(files: IDoc[]): Promise<IDocResponse[]> {

        var fileName: string;
        var file: IDoc = files.pop();

        if (!SPHelperCommon.isStringNullOrEmpty(file.fileName)) {
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

                    if (response.ok &&  !SPHelperCommon.isNull(response.result)) {

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
    public queryGETResquest(url: string): Promise<ISPBaseResponse> {
        return this.spQueryGET(url).then((response) => {
            return response;
        });
    }

    /**
     * Method gives access to query POST
     * @param postProps : ISPPostRequest object
     */
    public queryPOSTRequest(postProps: ISPPostRequest): Promise<ISPBaseResponse> {
        return this.spQueryPOST(postProps).then((response) => {
            return response;
        });
    }

    /**
     * Method gives access to query MERGE
     * @param postProps : ISPPostRequest object
     */
    public queryMERGERequest(postProps: ISPPostRequest): Promise<ISPBaseResponse> {
        return this.spQueryMERGE(postProps).then((response) => {
            return response;
        });
    }

    /**
     * Method gives access to query PATCH
     * @param postProps : ISPPostRequest object
     */
    public queryPATCHRequest(postProps: ISPPostRequest): Promise<ISPBaseResponse> {
        return this.spQueryPATCH(postProps).then((response) => {
            return response;
        });
    }
}

export { SPCommonOperations };