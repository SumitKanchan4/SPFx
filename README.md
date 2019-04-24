# SharePoint Framework Helper

[![SharePoint Framework Helper](/Images/sharePointWidgetsBanner.png?raw=true "SharePoint Framework Helper" )](https://www.sharepointwidgets.com)

In the need to develop the solution faster and easier I have created a [npm package](https://www.npmjs.com/package/spfxhelper), which will give you the flexibility to enjoy the much used logic handy and easy when writing code for SharePoint framework client side webparts or SharePoint framework extensions.

This [npm package](https://www.npmjs.com/package/spfxhelper) contains the most commonly used logic or I should say the operations that are used almost in every SPFx client side webpart or SPFx extensions. Using this npm package you can create your webparts much more easily and faster. This is very simple and handy.

This npm package code is also available on [GitHub](https://github.com/SumitKanchan4/SPFx) so if you are interested you can have a look at the implementation as well, or even you can request for the addition of new methods that you feel might help others as well.

# Why to use
This library is broken into different operations, so you only need to import the operations that you want to use. Below are the operations description:

### SPHelperBase 
>This is the base class and one cannot create object of this class. This class is responsible for the main interaction with SharePoint. This class is used internally from all the other classes. This class contains all the basic interaction methods with SharePoint.

### SPListOperations 
>This is the main class for all the operations related to SharePoint lists. It has methods for all the operations like getListByTitle, getListItemsByQuery, createListItem etc..

### SPFieldOperations
>This is the class responsible for the operations related to SharePoint Fields. It has methods like getFieldByList, addSiteColumnToList, createSiteColumn etc..

### SPCommonOperations
>This class is provided to give the flexible to the developers/users to get the information for some queries for which method are not found in any of the classes.

### SPHelperCommon
>This class contains only static methods and does not interact with the SharePoint. It contains some of the most common helper methods like isNull, stringIsNullOrEmpty etc..
### SPLogger
>This class can be used for the logging purpose. For using this you need to make some configuration manually
Create a SharePoint List with the title: ‘Error Logs’
Create Single line of text columns with the title
Error Type (Choice field- ERROR, DEBUG, INFO)
Error Description (Multiple line of text)

### SPBatchOperations
>This class gives the flexibility to use the batch operations in SharePoint Framework.

# Installation
To install the package you just need to write the following command in the node console

`npm install spfxhelper`


And your solution is ready for using the library. :-)

# How to use this library
The library is very simple to use. Below is the sample code how you can use

Import the spfxhelper in you code file so you can have all the available functionalities
```sh
import { SPListOperations } from 'spfxhelper';
```

Within the curly braces you can have one of the following Operations depending upon you need. For example purpose i’ll use SPListOperations. 

Now to use the methods, you cannot create object of the class, as the singleton approach has been adopted. To get the object of the class you need to get the instance of the object using getInstance method.

I’ll prefer to create the property so I don’t have to pass the parameters again and again.Below is the code sample

```sh
/**
  * This property returns the SPListOperation object
  */
 private get oListOperation(): SPListOperations {
   return SPListOperations.getInstance(this.context.spHttpClient as any, this.context.pageContext.web.absoluteUrl);
 }
```

Now you can use this property to access the methods.

# Methods
Below are the listing of all the methods available in the library

| Method Name | README |
| ------ | ------ |
| getInstance | returns the instance of the class |
|logError|logs the error in the SharePoint List|
|logDebug|logs the debug info in the SharePoint list|
|getListByTitle|returns the list by title|
|createList|creates the list|
|getListItemsByQuery|get the list items based on the query parameter|
|getListItemByID|get the list item based on the item ID|
|getListItems|returns all the list items based on the row count. If row count is not provided will return max i.e.. 100|
|createListItem|creates the list itme|
|createFolderInDocLib|creates the folder in the document library|
|createFolderInList|creates the folder in the list|
|updateListItem|updates the list item|
|getListMetadata|returns the list metadata required to create the list|
|getListsDetailsByBaseTemplateID|Returns the lists based on the template ID|
|getContentTypesByList|returns the content types associated with the list|
|getViewsByList|return the view associated with the list|
|isStringNullOrEmpy|checks for the string null|
|isNull|check for the object null|
|getFieldInternalName|returns the internal name of the field (removes the _x0020_)|
|getParameterValue|returns the parameter value from URL|
|addFieldToView|adds the field to the view in the list|
|getFieldByView|returns the field details by view|
|getFieldByList|returns the field details from the list|
|addSiteColumnToList|adds the site column to the list|
|createSiteColumn|creates the site column|
|getFieldBySite|returns the fields details of the site column|
|getColumnMetadata|returns the column metadata required for the creation of the column|
|addColumnToList|add the column to the list|
|getFieldsByList|returns all the fields associated with the list|
|getFieldsByView|returns all the fields associated with view in a list|
|getDocIconByFiles|returns the icons for the corresponding files|
|queryGETResquest|method to query any custom query with 'GET' verb|
|queryPOSTRequest |method to query any custom query with 'POST' verb|
|queryMERGERequest|method to query any custom query with 'MERGE' verb|
|queryPATCHRequest|method to query any custom query with 'PATCH' verb|
|oSPBatch|property returns the SPHttpClientBatch object required to create batch query|
|getBatchGETRequest|adds the url to the current SPHttpClientBatch object and returns response|
|getBatchPOSTRequest|adds the url to the current  SPHttpClientBatch object and returns response|
|SPHttpClientResponseToSPBaseResponse||


>Do share how I can improve this library, so it can help all the SharePoint community




Happy Coding

Sumit Kanchan
