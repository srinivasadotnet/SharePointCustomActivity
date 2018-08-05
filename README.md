SharePointCustomActivity

SharePoint Custom Activities is a reusable UI Path component which can be used by any RPA developer to reduce the effort of SharePoint Automation.

## Activities

- Get List By Name
- Insert List Items
- Upload Document To Library

### Get List By Name

This activity can be used to read all data from a SharePoint list. List can any list Custom list or default list like Calendar List, Task List etc.

##### Input Parameters :

   *SharePointSiteUri*   :  <string type>SharePoint site URL 

   *UserName* : <string type> SharePoint site user name

   *Password* : <string type> SharePoint site password

   *ListName* : <string type> List name where we wanted to perform read operations.

##### Output Parameters:

   *ListItems* : <DataTable type> output result will be returned in DataTable format.



### Insert List Items

This activity can be used to insert new items in a SharePoint list. List can any list Custom list or default list like Calendar List, Task List etc.

##### Input Parameters :

   *SharePointSiteUri* :  <string type>SharePoint site URL 

   *UserName* : <string type> SharePoint site user name

   *Password* : <string type> SharePoint site password

   *ListName* : <string type> List name where we wanted to perform read operations.

   *ListItems* : <DataTable type> DataTable with input records.

##### Output Parameters:

   *ResultMessage* : <string type> output result message, it is capable of returning of some user friendly error messages as well.



### Upload Document To Library

This activity can be used to upload new document/documents/folders to SharePoint Document Library. This will work only on Document Library type.

##### Input Parameters :

   *SharePointSiteUri* :  <string type>SharePoint site URL 

   *UserName* : <string type> SharePoint site user name

   *Password* : <string type> SharePoint site password

   *ListName* : <string type> List name where we wanted to perform read operations.

   *IsMultiFileUpload* : <boolean type> determine whether the Filepath is single file upload or Folder upload

   *FilePath* : <string type> input file path, or Folder path this is based on IsMultiFileUpload  parameter

##### Output Parameters:

   *ResultMessage* : <string type> output result message, it is capable of returning of user friendly error messages as well.



Used SharePoint CSOM(Client Side Object Model) technique to perform above activities. The core class is capable of handling CRUD operations on any type of SharePoint List's. 



As of now created only above mentioned activities. I will make sure it will support maximum operations on SharePoint lists. Development will be continued.  
