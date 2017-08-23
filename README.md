# SharePoint Framework Helper


This is the package for reducing the development time which will contain all the basic methods to implement the SharePoint Framework solution. I have tried to cover most of the basic operations that we use in every solution. I will be writing more blogs to make you understand how these are implemented and how we can use them. I will continue adding more and more methods with time, so keep watching.

## Methods

**SPLogger** : This class is helpful in logging if any of the issues are found. For now this checks for the list in the working directory in sharepoint with the following information
> List Name: Error Logs

> Required Columns: Title (default), Error Type (CHoice: ERROR,DEBUG,INFO), Error Description (Multiple line of text)

Following methods are contained in this class:
```
* getInstance   : returns the instance of the class
* logError      : logs the error in the SharePoint List
* logDebug      : logs the debug info in the sharepoint list
```


**SPListOperations** : This class is responsible for all the list based operations.

Following are the methods contained in this class
```
* getInstance           : returns the instance of the class
* getListByTitle        : returns the list by title
* createList            : creates the list
* getListItemsByQuery   : get the list items based on the query parameter
* getListItemByID       : get the list item based on the item ID
* getListItems          : returns all the list items based on the row count. If row count is not provided will return max
* createListItem        : creates the list itme
* createFolderInDocLib  : creates the folder in the document library
* createFolderInList    : creates the folder in the list
* updateListItem        : updates the list item
* getListMetadata       : returns the list metadata required to create the list
* getListsDetailsByBaseTemplateID   : Returns the lists based on the template ID
* getContentTypesByList : returns the content types associated with the list
* getViewsByList        : return the view associated with the list
```

**SPHelperCommon** : This class contains the most common methods

Following are the methods contained in this class:
```
* isStringNullOrEmpy    : checks for the string null
* isNull                : check for the object null
* getFieldInternalName  : returns the internal name of the field (removes the _x0020_)
* getParameterValue     : returns the parameter value from URL
```

**SPFieldOperations** : This class is responsible for all the field based operations:

Following methods are contained in this class:
```
* getInstance           : returns the instance of the class to access the methods
* addFieldToView        : adds the field to the view in the list
* getFieldByView        : returns the field details by view
* getFieldByList        : returns the field details from the list
* addSiteColumnToList   : adds the site column to the list
* createSiteColumn      : creates the site column
* getFieldBySite        : returns the fields details of the site column
* getColumnMetadata     : returns the column metadata required for the creation of the column
* addColumnToList       : add the column to the list
* getFieldsByList       : returns all the fields associated with the list
* getFieldsByView       : returns all the fields associated with view in a list
```

**SPCommonOperations** : This class is responsible for all the common operations in sharepoint

Following are the methods contained in this class:
```
* getInstance           : returns the instance of the class to access the methods
* getDocIconByFiles     : returns the doc icons for the corresponding files
* queryGETResquest      : method to query any custom query with 'GET' verb
* queryPOSTRequest      : method to query any custom query with 'POST' verb
* queryMERGERequest     : method to query any custom query with 'MERGE' verb
* queryPATCHRequest     : method to query any custom query with 'PATCH' verb
```

**SPBatchOperations** : This class is responsible for batch operations in sharepoint Framework

Following are the methods contained in this class
```
* get oSPBatch
* getBatchGETRequest
* getBatchPOSTRequest
* SPHttpClientResponseToSPBaseResponse
```


**_Since I have tried to cover all the basic needs, so anyone feels any method or functionality to be included , please feel free to drop me a mail or suggest on git._**'
###### Happy Coding...!!! :+1:

###### - Sumit Kanchan
