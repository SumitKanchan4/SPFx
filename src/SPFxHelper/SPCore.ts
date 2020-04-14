
import { UrlQueryParameterCollection, Validate, Text } from '@microsoft/sp-core-library';

const CLASS_NAME: string = `SPCore`;
/**
 * This class contains all the static methods that can be used in common.
 * Do not add any method that contains some specific logic
 */
class SPCore  {

    /** Checks is the string is null or empty or undefined */
    public static isEmptyString(value: string): boolean {
        try {
            Validate.isNonemptyString(value, 'value');
            return false;
        }
        catch (error) {
            return true;
        }
    }

    /** Checks is the object is null or not */
    public static isNull(value: any): boolean {
        try {

            Validate.isNotNullOrUndefined(value, "value");
            return false;
        }
        catch (error) {
            return true;
        }
    }

    /** Method returns the internal name by replacing the space with _x0020_ */
    public static getFieldInternalName(fieldName: string): string {
        try {
            return Text.replaceAll(fieldName, " ", '_x0020_');
        }
        catch (error) {
            // Cannot have the logging as the method is static and cannot have resources to intiate logger
            throw error;
        }
    }

    /** Method returns the parameter value from url */
    public static getParameterValue(url: string, paramName: string): string {

        var queryURL = new UrlQueryParameterCollection(url);
        return queryURL.getValue(paramName);
    }

    /**
     * Returns the age from current date
     * @param dateString Date time string in ISO format
     */
    public static calculateAge(dateString: string): number {

        try {
            let paramAge = new Date(dateString);
            let ageDifMs: number = Date.now() - paramAge.getTime();
            let ageDate: Date = new Date(ageDifMs); // miliseconds from epoch
            let age: number = ageDate.getUTCFullYear() - 1970;
            age = (age > 0) ? age : 0;
            return age;
        }
        catch (e) {
            return -1;
        }
    }

    /**
     * Returns the local storage (browser) object
     */
    public static getLocalStorage(): Storage {

        if (!SPCore.isNull(typeof (Storage))) {
            return localStorage;
        }
        else {
            return undefined;
        }
    }    
}

export { SPCore };