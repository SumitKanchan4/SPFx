/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 1 May 2017
 * Description: This class will contain only the SPField specific methods
 */


import { SPBase } from './SPBase';
import { SPCore } from './SPCore';
import { SPHttpClient } from '@microsoft/sp-http';
import { IFieldPOST, FieldType, FieldScope, IFields, IField } from './Props/ISPFieldProps';
import { ISPBaseResponse } from './Props/ISPBaseProps';
import { Log } from '@microsoft/sp-core-library';

const CLASS_NAME: string = `SPFieldOperations`;
/**
 * This class will contain only the SPField specific methods
 */
class SPFieldOperations extends SPBase {

    constructor(spHttpClient: SPHttpClient, webUrl: string, logSource: string) {
        super(spHttpClient, webUrl, logSource);
    }

    /** Use this method to get the SPFieldOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string, logSource: string): SPFieldOperations {

        return new SPFieldOperations(spHttpClient, webUrl, logSource);
    }

    /**
     * Method adds respective Field in respective View of respective List
     * @param listTitle : Title of the list in which view exists
     * @param viewName : Name of the view
     * @param fieldName : internal name of the field
     */
    public async addFieldToView(listTitle: string, viewName: string, fieldName: string): Promise<ISPBaseResponse> {

        let result: ISPBaseResponse;
        try {

            // Check if the column is already added to the following view
            let isAdded: boolean = await this.isFieldInView(listTitle, viewName, fieldName);

            if (!isAdded) {
                let url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/views/getByTitle('${viewName}')/ViewFields/addviewfield('${fieldName}')`;

                result = await this.spQueryPOST({ body: undefined, url: url });
            }
        } catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.addFieldToView`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method checks if the respective Field is the part of the respective View in respective List
     * @param listTitle : Title of the list in which view exists
     * @param viewName : Name of the view
     * @param fieldName : Internal Name of the field
     */
    public async isFieldInView(listTitle: string, viewName: string, fieldName: string): Promise<boolean> {

        let result: boolean = false;
        try {

            let response: ISPBaseResponse = await this.getFieldsByView(listTitle, viewName);

            if (response.ok) {
                result = (response.result.Items as string[]).indexOf(fieldName) > -1;
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.isFieldInView`));
            Log.error(this.LogSource, error);
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Returns all the fiedls associated with the view
     * @param listTitle : title of the list
     * @param viewName : title of the view
     * @returns All fields schema 
     * @returns Array of all the field Internal Name
     */
    public async getFieldsByView(listTitle: string, viewName: string): Promise<IFields> {

        let result: IFields = { ok: false };

        try {
            let url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/views/getByTitle('${viewName}')/ViewFields/`;
            let response: ISPBaseResponse = await this.spQueryGET(url);
            result = { ok: response.ok, details: response.result, error: response.error };
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getFieldsByView`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
    * Returns all the fiedls associated with the list
    * @param listTitle : title of the list
    */
    public async getFieldsByList(listTitle: string): Promise<IFields> {

        let result: IFields;

        try {
            let url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/fields/`;
            let response: ISPBaseResponse = await this.spQueryGET(url);
            result = { details: !!response.result.value ? response.result.value : response.result, ok: response.ok, error: response.error };
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getFieldsByList`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method checks if the Field exists in the List.
     * If it exists then returns the details of the field
     * @param fieldTitle : Title of the Field
     * @param listName : Title of the list
     */
    public async getFieldByList(fieldTitle: string, listName: string): Promise<IField> {

        let result: IField;

        try {
            let url = `${this.WebUrl}/_api/web/lists/getByTitle('${listName}')/fields?&$filter=Title eq '${fieldTitle}'`;
            let response: ISPBaseResponse = await this.spQueryGET(url);

            if (response.ok && !!response.result.value && response.result.value.length > 0) {
                result = { ok: true, detail: response.result.value[0], error: undefined };
            }
            else {
                // Check if any error occured or the field does not found
                result = { ok: response.ok, error: response.ok ? new Error(`Could not find the field with the specified title '${fieldTitle}' in the list`) : response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getFieldByList`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method adds the Site column to the List.
     * @param fieldTitle : Title of the field
     * @param listTitle : Title of the list
     * @param fieldSchemaXML : Schema Xml of the field
     */
    public async addSiteColumnToList(fieldTitle: string, listTitle: string): Promise<IField> {

        let result: IField;
        try {
            // Get the filed info
            let field: IField = await this.getFieldBySite(fieldTitle);

            if (field.ok) {
                let url = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/fields/createfieldasxml`;

                var data = JSON.stringify({
                    "parameters": {
                        "__metadata": { "type": "SP.XmlSchemaFieldCreationInformation" },
                        "Options": 8,
                        "SchemaXml": field.detail.SchemaXml
                    }
                });

                let response: ISPBaseResponse = await this.spQueryPOST({ body: data, url: url });
                result = { ok: response.ok, error: response.error, detail: response.result };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.addSiteColumnToList`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error, detail: undefined };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Method creates the Site column
     * @param field : IFieldPost object to retrieve the information required
     */
    public async createSiteColumn(field: IFieldPOST): Promise<IField> {

        let result: IField;
        try {
            let postUrl: string = `${this.WebUrl}/_api/web/fields/${field.fieldType == FieldType.Lookup ? 'addfield' : ''}`;
            let response: ISPBaseResponse = await this.spQueryPOST({ body: this.getColumnMetadata(field, FieldScope.SiteColumn), url: postUrl });
            result = { ok: response.ok, error: response.error, detail: response.result };
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.createSiteColumn`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /**
     * Checks if the fiels exists in site and returns with data 
     * @param fieldTitle :Title of the field
     */
    public async getFieldBySite(fieldTitle: string): Promise<IField> {
        let result: IField;

        try {
            let url = `${this.WebUrl}/_api/web/fields?$filter=Title eq '${fieldTitle}'`;
            let response: ISPBaseResponse = await this.spQueryGET(url);

            if (response.ok && !!response.result.value && response.result.value.length > 0) {
                result = { ok: true, detail: response.result.value[0], error: undefined };
            }
            else {
                // Check if any error occured or the field does not found
                result = { ok: response.ok, error: response.ok ? new Error(`Could not find the field with the specified title '${fieldTitle}' in the site`) : response.error };
            }
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.getFieldBySite`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

    /** 
     * Method returns the metadata of the solumn based on the IFieldPOST object 
     * */
    public getColumnMetadata(fieldDetails: IFieldPOST, fieldScope: FieldScope): string {

        var metadata: string;
        fieldDetails.allowMultiValues = SPCore.isNull(fieldDetails.allowMultiValues) ? false : fieldDetails.allowMultiValues;

        switch (fieldDetails.fieldType) {
            case FieldType.MultiChoice:
                var multiChoices = `'${fieldDetails.choices.join(`','`)}'`;
                fieldDetails.defaultValue = SPCore.isEmptyString(fieldDetails.defaultValue) ? fieldDetails.choices[0] : fieldDetails.defaultValue;
                metadata = `{ '__metadata': { 'type': 'SP.FieldMultiChoice' }, 'FieldTypeKind': 15, 'DefaultValue': '${fieldDetails.defaultValue}', {COLUMNGROUP}, 'Required': ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}', 'Choices': { '__metadata': { 'type': 'Collection(Edm.String)' }, 'results': [${multiChoices}] }, 'EditFormat': 0 }`;
                break;
            case FieldType.Choice:
                var choices = `'${fieldDetails.choices.join(`','`)}'`;
                fieldDetails.defaultValue = SPCore.isEmptyString(fieldDetails.defaultValue) ? fieldDetails.choices[0] : fieldDetails.defaultValue;
                metadata = `{ '__metadata': { 'type': 'SP.FieldChoice' }, 'FieldTypeKind': 6, 'DefaultValue': '${fieldDetails.defaultValue}', {COLUMNGROUP} 'Required':  ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}', 'Choices': { '__metadata': { 'type': 'Collection(Edm.String)' }, 'results': [${choices}] }, 'EditFormat': 0 }`;
                break;
            case FieldType.Number:
                metadata = `{ '__metadata': { 'type': 'SP.FieldNumber' }, 'FieldTypeKind': 9, {COLUMNGROUP} 'Required':  ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}'`;
                break;
            case FieldType.Lookup:
                metadata = `{ 'parameters': { '__metadata': { 'type': 'SP.FieldCreationInformation' }, 'FieldTypeKind': 7, {COLUMNGROUP} 'Required':  ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}', 'LookupListId': '${fieldDetails.lookupListID}', 'LookupFieldName': '${fieldDetails.lookupColName}' }}`;
                break;
            case FieldType.Note:
                metadata = `{ '__metadata': { 'type': 'SP.FieldMultiLineText' }, 'FieldTypeKind': 3,{COLUMNGROUP} 'Required':  ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}', 'NumberOfLines': 12, 'RichText': true, 'AllowHyperlink' : true, 'RestrictedMode': true}`;
                break;
            case FieldType.Text:
                metadata = `{ '__metadata': { 'type': 'SP.Field' }, 'FieldTypeKind': 2, {COLUMNGROUP} 'Required':  ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}'}`;
                break;
            case FieldType.User:
                metadata = `{ '__metadata': { 'type': 'SP.FieldUser' }, 'FieldTypeKind': 20, {COLUMNGROUP} 'Required':  ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}', 'SelectionMode': 0, 'Presence': true}`;
                break;
            case FieldType.Boolean:
                break;
            case FieldType.DateTime:
                break;
        }

        metadata = fieldScope == FieldScope.ListColumn ? metadata.replace('{COLUMNGROUP}', '') : metadata.replace('{COLUMNGROUP}', `'Group': '${fieldDetails.group}',`);

        return metadata;
    }

    /**
     * Method creates the column to the list
     * @param field : IFieldPOST object to retrieve required information
     */
    public async createListColumn(field: IFieldPOST): Promise<IField> {

        let result: IField;
        try {
            let postUrl: string = `${this.WebUrl}/_api/web/lists/getByTitle('${field.listName}')/fields/${field.fieldType == FieldType.Lookup ? 'addfield' : ''}`;
            let response: ISPBaseResponse = await this.spQueryPOST({ body: this.getColumnMetadata(field, FieldScope.ListColumn), url: postUrl });

            // check if the column needs to be added to the view, then view name and list name are manadatory
            if (field.addToView && !SPCore.isEmptyString(field.viewName) && !SPCore.isEmptyString(field.listName))
                await this.addFieldToView(field.listName, field.viewName, response.result.InternalName);

            result = { ok: response.ok, error: response.error, detail: response.result };
        }
        catch (error) {
            Log.error(this.LogSource, new Error(`Error occured in ${CLASS_NAME}.createListColumn`));
            Log.error(this.LogSource, error);
            result = { ok: false, error: error };
        }
        finally {
            return Promise.resolve(result);
        }
    }

}

export { SPFieldOperations };