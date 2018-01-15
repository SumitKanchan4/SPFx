import { SPListOperations } from './SPListOperations';
import { SPFieldOperations } from './SPFieldOperations';
import { SPHelperCommon } from './SPHelperCommon';
import { SPHttpClient } from '@microsoft/sp-http';
import { errorType, ILogger, ILoggerResponse } from './Props/ISPLogProps';
import { IListGET } from './Props/ISPListProps';
import { IFieldGET } from './Props/ISPFieldProps';

/**
 * Generic Logger
 */
class SPLogger {

    private static instance: SPLogger;
    private spListOperation: SPListOperations;
    private spFieldOperation: SPFieldOperations;

    private writeDebug: boolean;    // Notifies when to write the debug logs
    private tryCount: number = 0;   // notifies the number of attempt to validate the error list structure
    private maxTryCount: number = 3;    // max try count to validate the error list structure
    private isValid: boolean = false;   // notifies if the structure is valid
    private errorOccured: boolean = false;  // notifies if the error occured while writing error

    private lstErrorLog: string = 'Error Logs';
    private colErrorType: string = 'Error Type';
    private colError: string = 'Error Description';
    private errorStatus: number = -1;

    private constructor(spHttpClient: SPHttpClient, webUrl: string, writeDebug: boolean) {
        this.writeDebug = writeDebug;
        this.spListOperation = SPListOperations.getInstance(spHttpClient, webUrl);
        this.spFieldOperation = SPFieldOperations.getInstance(spHttpClient, webUrl);
    }

    /** Use this method to get the SPLogger class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string, writeDebug: boolean): SPLogger {

        SPLogger.instance = SPHelperCommon.isNull(SPLogger.instance) ? new SPLogger(spHttpClient, webUrl, writeDebug) : SPLogger.instance;
        SPLogger.instance.writeDebug = writeDebug;
        return SPLogger.instance;
    }

    /**
     * Method writes the ERROR information in the SharePoint List 'Error Logs'
     * @param logInfo : ILogger object to get error info
     */
    public logError(logInfo: ILogger): Promise<ILoggerResponse> {

        return this.writeToLogList(logInfo.errorMethod, parseInt(errorType.ERROR.toString()), logInfo.errorMessage).then((resp) => {
            return Promise.resolve(resp);
        });
    }

    /**
     * Method writes the DEBUG information in the SharePoint List 'Error Logs'.
     * Will only write if the writeDebug Property is true
     * @param logInfo : ILogger object to get error info
     */
    public logDebug(logInfo: ILogger): Promise<ILoggerResponse> {

        if (this.writeDebug) {
            return this.writeToLogList(logInfo.errorMethod, parseInt(errorType.DEBUG.toString()), logInfo.errorMessage).then((resp) => {
                return Promise.resolve(resp);
            });
        }
    }

    /**
     * Writes the error in the SharePoint list.
     * @param errorMethod : class and method name in which error occured
     * @param errorTypeIndex : error type index
     * @param errorMessage : error message
     */
    private writeToLogList(errorMethod: string, errorTypeIndex: number, errorMessage: string): Promise<ILoggerResponse> {
        try {

            return this.validateLogList().then((validResp) => {

                if (validResp.success && !this.errorOccured) {

                    return this.spListOperation.getListByTitle(this.lstErrorLog).then((lstResp) => {

                        var body: string = JSON.stringify({
                            '__metadata': {
                                'type': `${lstResp.details.ListItemEntityTypeFullName}`
                            },
                            'Title': `${errorMethod}`,
                            'Error_x0020_Type': `${errorType[errorTypeIndex]}`,
                            'Error_x0020_Description': `${errorMessage}`
                        });

                        return this.spListOperation.createListItem(this.lstErrorLog, body).then((creteResp) => {

                            var loggerResponse: ILoggerResponse = {
                                errorMethod: creteResp.errorMethod,
                                ok: creteResp.ok,
                                status: creteResp.status,
                                statusText: creteResp.statusText,
                                success: true
                            };
                            return Promise.resolve(loggerResponse);
                        });
                    });
                }
            });
        }
        catch (error) {
            this.errorOccured = true;
            return Promise.resolve({
                errorMethod: 'SPLogger.writeToLogList',
                ok: false,
                status: this.errorStatus,
                statusText: error.message,
                success: false
            });
        }
    }

    /**
     * Method validates the error list structure for the required columns.
     * This will try for max 3 count.
     */
    private validateLogList(): Promise<ILoggerResponse> {

        var logResponse: ILoggerResponse;

        if (!this.isValid && this.tryCount <= this.maxTryCount) {

            //1. Check if the list exists
            return this.spListOperation.getListByTitle(this.lstErrorLog).then((lstresp) => {

                if (!lstresp.ok) {
                    this.isValid = false;
                    this.tryCount += 1;
                    return this.typeCastListToLog(lstresp);
                }
                else if (lstresp.ok && lstresp.exists) {

                    //2. CHeck if the columns exists
                    return this.spFieldOperation.getFieldByList(this.colError, this.lstErrorLog).then((colErrorResp) => {

                        if ((!colErrorResp.ok) || (colErrorResp.ok && !colErrorResp.exists)) {
                            this.isValid = false;
                            this.tryCount += 1;
                            return Promise.resolve(this.typeCastFieldToLog(colErrorResp));
                        }
                        else {
                            return this.spFieldOperation.getFieldByList(this.colErrorType, this.lstErrorLog).then((colETResp) => {

                                // Recieves response as object if exists else recieves boolean as false
                                if ((!colETResp.ok) || (colETResp.ok && !colETResp.exists)) {
                                    this.isValid = false;
                                    this.tryCount += 1;
                                    return Promise.resolve(this.typeCastFieldToLog(colETResp));
                                }
                                else {
                                    this.isValid = true;
                                    return Promise.resolve(this.typeCastFieldToLog(colETResp));
                                }
                            });
                        }
                    });
                }
                else {
                    return Promise.resolve({
                        errorMethod: 'SPLogger.validateLogList',
                        ok: false,
                        status: this.errorStatus,
                        statusText: `Could not find the list: ${this.lstErrorLog} `,
                        success: false
                    });
                }
            });
        }
        else if (!this.isValid && this.tryCount > this.maxTryCount) {
            return Promise.resolve({
                errorMethod: 'SPLogger.validateLogList',
                ok: false,
                status: this.errorStatus,
                statusText: `Try writing count exceeded its max limit of ${this.maxTryCount} `,
                success: false
            });
        }
        else {
            return Promise.resolve({
                errorMethod: 'SPLogger.validateLogList',
                ok: true,
                status: 0,
                statusText: ``,
                success: true
            });
        }
    }

    /**
     * Method typecasts the IListGet object to ILoggerResponse.
     * @param listResponse : IListGet object
     * @param errorMethod : errorMethod if needs to be updated
     */
    private typeCastListToLog(listResponse: IListGET, errorMethod?: string): ILoggerResponse {

        var loggerResponse: ILoggerResponse = {
            errorMethod: SPHelperCommon.isStringNullOrEmpty(errorMethod) ? listResponse.errorMethod : errorMethod,
            ok: listResponse.ok,
            status: listResponse.status,
            statusText: listResponse.statusText,
            success: listResponse.exists
        };
        return loggerResponse;
    }

    /**
     * Method typecasts the IFieldGET object to ILoggerResponse.
     * @param listResponse : IListGet object
     * @param errorMethod : errorMethod if needs to be updated
     */
    private typeCastFieldToLog(listResponse: IFieldGET, errorMethod?: string): ILoggerResponse {

        var loggerResponse: ILoggerResponse = {
            errorMethod: SPHelperCommon.isStringNullOrEmpty(errorMethod) ? listResponse.errorMethod : errorMethod,
            ok: listResponse.ok,
            status: listResponse.status,
            statusText: listResponse.statusText,
            success: listResponse.exists
        };
        return loggerResponse;
    }
}

export { SPLogger };