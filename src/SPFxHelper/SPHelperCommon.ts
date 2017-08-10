
import { UrlQueryParameterCollection, Validate } from '@microsoft/sp-core-library';

/**
 * This class contains all the static methods that can be used in common.
 * Do not add any method that contains some specific logic
 */
class SPHelperCommon {

    /** Checks is the string is null or empty or undefined */
    public static isStringNullOrEmpy(value: string): boolean {
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

            var internalName: string = fieldName.replace(/ /g, '_x0020_');
            return internalName;
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
}

export { SPHelperCommon };