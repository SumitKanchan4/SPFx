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
* getInstance       : returns the instance of the class
* getListByTitle    : returns the list by title
* createList   
* getListItemsByQuery
* getListItemByID
* getListItems
* createListItem
* createFolderInDocLib
* createFolderInList
* updateListItem
* getListMetadata
* getListsDetailsByBaseTemplateID
```

**SPHelperCommon** : This class contains the most common methods

Following are the methods contained in this class:
```
* isStringNullOrEmpy
* isNull
* getFieldInternalName
* getParameterValue
```

**SPFieldOperations** : This class is responsible for all the field based operations:

Following methods are contained in this class:
```
* getInstance
* addFieldToView
* getFieldByView
* getFieldByList
* addSiteColumnToList
* createSiteColumn
* getFieldBySite
* getColumnMetadata
* addColumnToList
```

**SPCommonOperations** : This class is responsible for all the common operations in sharepoint

Following are the methods contained in this class:
```
* getInstance
* getDocIconByFiles
* queryGETResquest
* queryPOSTRequest
* queryMERGERequest
* queryPATCHRequest
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
