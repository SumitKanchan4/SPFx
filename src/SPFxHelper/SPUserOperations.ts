import { SPBase } from './SPBase';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPCore } from './SPCore';

class SPUserOperations extends SPBase {

    private static webUrl: string = undefined;

    constructor(spHttpClient: SPHttpClient, webUrl: string, logSource: string) {
        super(spHttpClient, webUrl, logSource);
    }

    /** Use this method to get the SPCommonOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string, logSource: string): SPUserOperations {

        return new SPUserOperations(spHttpClient, webUrl, logSource);
    }


    private static isUserGroup(groupName: string, email: string): void {

        let query: string = `${this.webUrl}/_api/web/SiteGroup/GetByName('${groupName}')?$filter=Email eq '${email}'`;

    }
}

export { SPUserOperations };