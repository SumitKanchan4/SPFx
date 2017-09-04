import { SPHelperBase } from './SPHelperBase';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPHelperCommon } from './SPHelperCommon';

class SPUserOperations extends SPHelperBase {

    private static instance: SPUserOperations;

    private constructor(spHttpClient: SPHttpClient, webUrl: string) {
        super(spHttpClient, webUrl);
    }

    /** Use this method to get the SPUserOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string): SPUserOperations {

        SPUserOperations.instance = SPHelperCommon.isNull(SPUserOperations.instance) ? new SPUserOperations(spHttpClient, webUrl) : SPUserOperations.instance;

        return SPUserOperations.instance;

    }

    
    public static isUserGroup(groupName: string, email : string):void{



    }
}

export { SPUserOperations };