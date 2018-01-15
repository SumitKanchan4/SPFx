import { SPHelperBase } from './SPHelperBase';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPHelperCommon } from './SPHelperCommon';

class SPUserOperations extends SPHelperBase {

    private static instance: SPUserOperations;
    private static webUrl: string = undefined;

    private constructor(spHttpClient: SPHttpClient, webUrl: string) {
        super(spHttpClient, webUrl);
    }

    /** Use this method to get the SPUserOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string): SPUserOperations {

        SPUserOperations.instance = SPHelperCommon.isNull(SPUserOperations.instance) ? new SPUserOperations(spHttpClient, webUrl) : SPUserOperations.instance;
        SPUserOperations.webUrl = webUrl;
        return SPUserOperations.instance;

    }


    public static isUserGroup(groupName: string, email: string): void {

        let query: string = `${this.webUrl}/_api/web/SiteGroup/GetByName('${groupName}')?$filter=Email eq '${email}'`;

    }
}

export { SPUserOperations };