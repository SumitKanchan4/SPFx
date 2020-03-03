# **SharePoint Framework Helper**


[![SharePoint Framework Helper](https://github.com/SumitKanchan4/SPFx/raw/master/Images/sharePointWidgetsBanner.png?raw=true "SharePoint Framework Helper" )](https://www.sharepointwidgets.com)

In the need to develop the solution faster and easier I have created a [npm package](https://www.npmjs.com/package/spfxhelper), which will give you the flexibility to enjoy the much used logic handy and easy when writing code for SharePoint framework client side webparts or SharePoint framework extensions.

This [npm package](https://www.npmjs.com/package/spfxhelper) contains the most commonly used functions or I should say the operations that are used almost in every SPFx client side webpart or SPFx extensions. Using this npm package you can create your webparts much more easily and faster. This is very simple and handy.

This npm package code is also available on [GitHub](https://github.com/SumitKanchan4/SPFx) so if you are interested you can have a look at the implementation as well, or even you can request for the addition of new methods that you feel might help others as well.

## **Why to use ?**
- This library contains handy functions that will save lot of man hours in implementing
- Well designed results so it will help in writing the code efficiently
- Well broken into the type of operation, so include only the class that is most useful to you

## **How to use ?**
To install the package execute the below command
```sh
> npm i spfxhelper
```
And then import the following in the code file
```sh
import { "CLASS_NAME" } from 'spfxhelper';
```

## **Inside the library**
- This library contains a number of classes

### [`SPListOperations`](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations)
>*This class contains all the methods that are responsible for the interaction of the SharePoint List.*

| Methods | Description |
| ------ | ------ |
| [**createFolderInDocLib**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#createFolderInDocLib) | creates the folder in document library |
| [**createFolderInList**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#createFolderInList) | creates the folder in list |
| [**createList**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#createList) | creates the list |
| [**createListItem**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#createListItem) | creates the list item |
| [**getContentTypesByList**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getContentTypesByList) | Returns the content types associated with the list |
| [**getInstance**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getinstance) | returns the instance of the class |
| [**getListBytitle**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getlistbytitle) | returns the list details by title |
| [**getListItemByID**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getListItemByID) | get the list items based on the item ID |
| [**getListItems**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getListItems) | returns all the list items based on the row count. If row count is not provided will return max i.e.. 100 |
| [**getListItemsByQuery**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getListItemsByQuery) | get the list items based on the query parameter |
| [**getListMetadata**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getListMetadata) | returns the list metadata required to create the list |
| [**getListsDetailsByBaseTemplateID**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getListsDetailsByBaseTemplateID) | Returns the lists based on the template ID |
| [**getViewsByList**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#getViewsByList) | Returns the lists views |
| [**updateListItem**](https://github.com/SumitKanchan4/SPFx/wiki/List-Operations#updateListItem) | updates the list item |

***
### [`SPFieldOperations`](https://github.com/SumitKanchan4/SPFx/wiki/Field-Operations)
>*This is the class responsible for the operations related to SharePoint Fields.*

| Methods | Description |
| ------ | ------ |
| [**addColumnToList**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#addColumnToList) | add the column to the list |
| [**addFieldToView**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#addFieldToView) | adds the field to the view in the list |
| [**addSiteColumnToList**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#addSiteColumnToList) | adds the site column to the list |
| [**createListColumn**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#createListColumn) | creates the list column |
| [**createSiteColumn**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#createSiteColumn) | creates the site column |
| [**getColumnMetadata**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getColumnMetadata) | returns the column metadata required for the creation of the column |
| [**getFieldByList**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getFieldByList) | returns the field details from the list |
| [**getFieldBySite**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getFieldBySite) | returns the fields details of the site column |
| [**getFieldByView**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getFieldByView) | returns the field details by view |
| [**getFieldsByList**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getFieldsByList) | returns all the fields associated with the list |
| [**getFieldsByView**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getFieldsByView) | returns all the fields associated with view in a list |

***
### [`SPCommonOperations`](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations)
>*This class is the extension of the base class where one can query the API's which are not available as methods in other classes*

| Methods | Description |
| ------ | ------ |
| [**getDocIconByFiles**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getDocIconByFiles) | returns the OOB icons for the file types |
| [**getInstance**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#getinstance) | returns the instance of the class |
| [**queryGETResquest**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#queryGETResquest) | method to query any custom query with 'GET' verb |
| [**queryMERGERequest**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#queryMERGERequest) | method to query any custom query with 'MERGE' verb |
| [**queryPATCHRequest**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#queryPATCHRequest) | method to query any custom query with 'PATCH' verb |
| [**queryPOSTRequest**](https://github.com/SumitKanchan4/SPFx/wiki/Common-Operations#queryPOSTRequest) | method to query any custom query with 'POST' verb |

***
### [`SPCore`](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations) (renamed from SPHelperCommon)
>*This class contains only static methods and contains the helper methods*

| Methods | Description |
| ------ | ------ |
| [**calculateAge**](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations#calculateAge) | returns the age from the specified date |
| [**getFieldInternalName**](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations#getFieldInternalName) | returns the internal name of the field (removes the _x0020_) |
| [**getLocalStorage**](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations#getLocalStorage) | returns the local storage object |
| [**getParameterValue**](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations#getParameterValue) | returns the parameter value from URL |
| [**isEmptyString**](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations#isemptystring) | checks for the string emptiness |
| [**isNull**](https://github.com/SumitKanchan4/SPFx/wiki/Core-Operations#isNull) | check for the object null |
***
### SPLogger
>*This class has been removed from the library. Instead use the OOB class [Log](https://docs.microsoft.com/en-us/javascript/api/sp-core-library/log?view=sp-typescript-latest) in sp-core-library.*

```sh
import { Log } from '@microsoft/sp-core-library';
```
Use the following methods
```sh
static error(source: string, error: Error, scope?: ServiceScope): void;
static info(source: string, message: string, scope?: ServiceScope): void;
static verbose(source: string, message: string, scope?: ServiceScope): void;
static warn(source: string, message: string, scope?: ServiceScope): void;
```
***

> Do share how I can improve this library, so it can help all the SharePoint community




Happy Coding

*#SharePointWidgets #MicrosoftSharePoint*

**Sumit Kanchan**
