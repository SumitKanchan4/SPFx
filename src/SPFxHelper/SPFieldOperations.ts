/**
 * Created By: Sumit Kanchan
 * Created on: 1 May 2017
 * Modified By: Sumit Kanchan
 * Modified on: 1 May 2017
 * Description: This class will contain only the SPField specific methods
 */


import { SPHelperBase } from './SPHelperBase';
import { SPHelperCommon } from './SPHelperCommon';
import { SPHttpClient } from '@microsoft/sp-http';
import { IFieldGET, IFieldPOST, FieldType, FieldScope } from './Props/ISPFieldProps';

/**
 * This class will contain only the SPField specific methods
 */
class SPFieldOperations extends SPHelperBase {

    private static instance: SPFieldOperations;

    constructor(spHttpClient: SPHttpClient, webUrl: string, ) {
        super(spHttpClient, webUrl);
    }

    /** Use this method to get the SPFieldOperations class Object */
    public static getInstance(spHttpClient: SPHttpClient, webUrl: string): SPFieldOperations {

        SPFieldOperations.instance = SPHelperCommon.isNull(SPFieldOperations.instance) ? new SPFieldOperations(spHttpClient, webUrl) : SPFieldOperations.instance;

        return SPFieldOperations.instance;

    }

    /**
     * Method adds respective Field in respective View of respective List
     * @param listTitle : Title of the list in which view exists
     * @param viewName : Name of the view
     * @param fieldTitle : Title of the field
     */
    public addFieldToView(listTitle: string, viewName: string, fieldTitle: string): Promise<IFieldGET> {
        try {

            // Check if the column is already added to the following view
            return this.getFieldByView(listTitle, viewName, fieldTitle).then((existResp) => {

                if (existResp.ok && !existResp.exists) {

                    var url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/views/getByTitle('${viewName}')/ViewFields/addviewfield('${fieldTitle}')`;

                    // IF the field does not exist than add the field to the view
                    return this.spQueryPOST({
                        body: null,
                        url: url
                    }).then((response) => {

                        var field: IFieldGET = {
                            exists: false,
                            id: null,
                            ok: response.ok,
                            schemaXml: null,
                            status: response.status,
                            statusText: response.statusText,
                            title: fieldTitle,
                            errorMethod: 'SPFieldOperations.addColumnToView'
                        };

                        return Promise.resolve(field);
                    });
                } else {
                    return Promise.resolve(existResp);
                }
            });

        } catch (error) {
            Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: fieldTitle,
                errorMethod: 'SPFieldOperations.addColumnToView'
            });
        }
    }

    /**
     * Method checks if the respective Field is the part of the respective View in respective List
     * @param listTitle : Title of the list in which view exists
     * @param viewName : Name of the view
     * @param fieldTitle : Title of the field
     */
    public getFieldByView(listTitle: string, viewName: string, fieldTitle: string): Promise<IFieldGET> {
        try {
            var url: string = `${this.WebUrl}/_api/web/lists/getByTitle('${listTitle}')/views/getByTitle('${viewName}')/ViewFields/`;

            return this.spQueryGET(url).then((response) => {

                var field: IFieldGET = {
                    exists: false,
                    id: null,
                    ok: response.ok,
                    schemaXml: null,
                    status: response.status,
                    statusText: response.statusText,
                    title: fieldTitle,
                    errorMethod: 'SPFieldOperations.fieldExistsInView'
                };

                if (response.ok && response.result.value.length > 0) {

                    var fieldInternalName = SPHelperCommon.getFieldInternalName(fieldTitle);

                    response.result.value.Items.forEach(element => {

                        if (element == fieldInternalName || element == fieldTitle) {
                            field.exists = true;
                        }
                    });
                }

                return Promise.resolve(field);
            });
        } catch (error) {
            Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: fieldTitle,
                errorMethod: 'SPFieldOperations.fieldExistsInView'
            });
        }
    }

    /**
     * Method checks if the Field exists in the List.
     * If it exists then returns the details of the field
     * @param fieldTitle : Title of the Field
     * @param lstName : Title of the list
     */
    public getFieldByList(fieldTitle: string, lstName: string): Promise<IFieldGET> {

        try {
            var url = `${this.WebUrl}/_api/web/lists/getByTitle('${lstName}')/fields?$select=Title,SchemaXml,Id&$filter=Title eq '${fieldTitle}'&$top=1`;

            return this.spQueryGET(url).then((response) => {

                var field: IFieldGET;

                if (response.ok && response.result.value.length > 0) {
                    field = {
                        exists: true,
                        id: response.result.value.Id,
                        ok: response.ok,
                        schemaXml: response.result.value.SchemaXml,
                        status: response.status,
                        statusText: response.statusText,
                        title: fieldTitle,
                        errorMethod: 'SPFieldOperations.checkListColExists'
                    };
                }
                else {
                    field = {
                        exists: false,
                        id: null,
                        ok: response.ok,
                        schemaXml: null,
                        status: response.status,
                        statusText: response.statusText,
                        title: fieldTitle,
                        errorMethod: response.errorMethod
                    };
                }
                return Promise.resolve(field);
            });
        } catch (error) {
            Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: fieldTitle,
                errorMethod: 'SPFieldOperations.checkListColExists'
            });
        }
    }

    /**
     * Method adds the Site column to the List.
     * @param fieldTitle : Title of the field
     * @param listTitle : Title of the list
     * @param fieldSchemaXML : Schema Xml of the field
     */
    public addSiteColumnToList(fieldTitle: string, listTitle: string, fieldSchemaXML: string): Promise<IFieldGET> {

        var fieldDetails: IFieldGET;
        try {

            return this.getFieldByList(fieldTitle, listTitle).then((existResp) => {

                if ((existResp.ok && existResp.exists) || !existResp.ok) {
                    return existResp;
                }
                else {

                    var url = `${this.WebUrl}/_api/web/lists/getByTitle('${fieldTitle}')/fields/createfieldasxml`;

                    var data = JSON.stringify({
                        "parameters": {
                            "__metadata": { "type": "SP.XmlSchemaFieldCreationInformation" },
                            "Options": 8,
                            "SchemaXml": fieldSchemaXML
                        }
                    });

                    return this.spQueryPOST({ body: data, url: url }).then((createResp) => {

                        fieldDetails = {
                            exists: true,
                            id: createResp.result.value.Id,
                            ok: createResp.ok,
                            schemaXml: createResp.result.value.SchemaXml,
                            status: createResp.status,
                            statusText: createResp.statusText,
                            title: fieldTitle,
                            errorMethod: 'SPFieldOperations.addSiteColumnToList'
                        };
                        return Promise.resolve(fieldDetails);
                    });
                }
            });
        } catch (error) {
            return Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: fieldTitle,
                errorMethod: 'SPFieldOperations.addSiteColumnToList'
            });
        }
    }

    /**
     * Method creates the Site column
     * @param field : IFieldPost object to retrieve the information required
     */
    public createSiteColumn(field: IFieldPOST): Promise<IFieldGET> {

        var postUrl: string = `${this.WebUrl}/_api/web/fields`;
        postUrl = field.fieldType == FieldType.Lookup ? `${postUrl}/addfield` : postUrl;

        try {
            if (!SPHelperCommon.isNull(field)) {

                return this.getFieldBySite(field.title).then((colResponse) => {

                    // If the column already exists or the response fails
                    if ((colResponse.ok && colResponse.exists) || !colResponse.ok) {
                        return colResponse;
                    }
                    // If the column does not exists
                    else if (colResponse && !colResponse.exists) {

                        return this.spQueryPOST({ url: postUrl, body: this.getColumnMetadata(field, FieldScope.SiteColumn) }).
                            then((response) => {

                                var fieldDetails: IFieldGET = {
                                    exists: true,
                                    id: response.result.value.Id,
                                    ok: response.ok,
                                    schemaXml: response.result.value.SchemaXml,
                                    status: response.status,
                                    statusText: response.statusText,
                                    title: field.title,
                                    errorMethod: 'SPFieldOperations.createSiteColumn'
                                };
                                return Promise.resolve(fieldDetails);
                            });
                    }
                });
            }
            else {
                var fieldDetails: IFieldGET = {
                    exists: false,
                    id: null,
                    ok: false,
                    schemaXml: null,
                    status: this.errorStatus,
                    statusText: 'Parameter object cannot be null',
                    title: field.title,
                    errorMethod: 'SPFieldOperations.createSiteColumn'
                };
                return Promise.resolve(fieldDetails);
            }
        }
        catch (error) {
            return Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: field.title,
                errorMethod: 'SPFieldOperations.createSiteColumn'
            });
        }
    }

    /**
     * Checks if the fiels exists in site and returns with data 
     * @param fieldTitle :Title of the field
     */
    public getFieldBySite(fieldTitle: string): Promise<IFieldGET> {
        var fieldDetails: IFieldGET;

        try {

            var url = `${this.WebUrl}/_api/web/fields?$filter=Title eq '${fieldTitle}'&$top=1`;

            return this.spQueryGET(url).then((response) => {

                if (response.ok && response.result.value.length > 0) {

                    fieldDetails = {
                        exists: true,
                        id: response.result.value[0].Id,
                        ok: response.ok,
                        schemaXml: response.result.value[0].SchemaXml,
                        status: response.status,
                        statusText: response.statusText,
                        title: fieldTitle,
                        errorMethod: 'SPFieldOperations.fieldExistsInSite'
                    };
                }
                else {
                    fieldDetails = {
                        exists: false,
                        id: null,
                        ok: response.ok,
                        schemaXml: null,
                        status: response.status,
                        statusText: response.statusText,
                        title: fieldTitle,
                        errorMethod: response.errorMethod
                    };
                }

                return Promise.resolve(fieldDetails);
            });
        }
        catch (error) {

            return Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: fieldTitle,
                errorMethod: 'SPFieldOperations.fieldExistsInSite'
            });
        }
    }

    /** Method returns the metadata of the solumn based on the IFieldPOST object */
    public getColumnMetadata(fieldDetails: IFieldPOST, fieldScope: FieldScope): string {

        var metadata: string;
        fieldDetails.allowMultiValues = SPHelperCommon.isNull(fieldDetails.allowMultiValues) ? false : fieldDetails.allowMultiValues;

        switch (fieldDetails.fieldType) {
            case FieldType.MultiChoice:
                var multiChoices = `'${fieldDetails.choices.join(`','`)}'`;
                fieldDetails.defaultValue = SPHelperCommon.isStringNullOrEmpy(fieldDetails.defaultValue) ? fieldDetails.choices[0] : fieldDetails.defaultValue;
                metadata = `{ '__metadata': { 'type': 'SP.FieldMultiChoice' }, 'FieldTypeKind': 15, 'DefaultValue': '${fieldDetails.defaultValue}', {COLUMNGROUP}, 'Required': ${fieldDetails.isRequired}, 'Title': '${fieldDetails.title}', 'Choices': { '__metadata': { 'type': 'Collection(Edm.String)' }, 'results': [${multiChoices}] }, 'EditFormat': 0 }`;
                break;
            case FieldType.Choice:
                var choices = `'${fieldDetails.choices.join(`','`)}'`;
                fieldDetails.defaultValue = SPHelperCommon.isStringNullOrEmpy(fieldDetails.defaultValue) ? fieldDetails.choices[0] : fieldDetails.defaultValue;
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
     * Method adds the column to the list
     * @param field : IFieldPOST object to retrieve required information
     */
    public addColumnToList(field: IFieldPOST): Promise<IFieldGET> {

        var postUrl: string = `${this.WebUrl}/_api/web/lists/getByTitle('${field.listName}')/fields`;
        postUrl = field.fieldType == FieldType.Lookup ? `${postUrl}/addfield` : postUrl;
        var fieldDetails: IFieldGET;

        try {
            if (!SPHelperCommon.isNull(field)) {

                return this.getFieldByList(field.title, field.listName).then((existResp) => {

                    if ((existResp.ok && existResp.exists) || !existResp.ok) {
                        return existResp;
                    }
                    else if (existResp.ok && !existResp.exists) {

                        return this.spQueryPOST({ url: postUrl, body: this.getColumnMetadata(field, FieldScope.ListColumn) }).then((fieldResp) => {
                            fieldDetails = {
                                exists: true,
                                id: fieldResp.result.value[0].Id,
                                ok: fieldResp.ok,
                                schemaXml: fieldResp.result.value[0].SchemaXml,
                                status: fieldResp.status,
                                statusText: fieldResp.statusText,
                                title: field.title,
                                errorMethod: 'SPFieldOperations.addColumnToList'
                            };

                            return Promise.resolve(fieldDetails);

                        });
                    }
                });
            }
        }
        catch (error) {
            return Promise.resolve({
                exists: false,
                id: null,
                ok: false,
                schemaXml: null,
                status: this.errorStatus,
                statusText: error.message,
                title: field.title,
                errorMethod: 'SPFieldOperations.addColumnToList'
            });
        }

    }
}

export { SPFieldOperations };