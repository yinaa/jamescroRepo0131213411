# Use Microsoft Graph to work with Excel

## Accessing Excel workbooks and resources

You can use Microsoft Graph to allow web and mobile applications to read and modify Excel workbooks stored in OneDrive, SharePoint, or other supported storage platforms. The `Workbook` (or Excel file) resource contains all the other Excel resources through relationships. You can access a workbook through the [Drive API](../api-reference/v1.0/resources/drive.md) by identifying the location of the file in the URL. For example:

`https://graph.microsoft.com/{version}/me/drive/items/{id}/workbook/`  
`https://graph.microsoft.com/{version}/me/drive/root:/{item-path}:/workbook/`  

You can access a set of Excel objects (such as Table, Range, or Chart) by using standard REST APIs to perform  create, read, update, and delete (CRUD) operations on the workbook. For example, 
`GET https://graph.microsoft.com/{version}/me/drive/items/{id}/workbook/worksheets`  
returns a collection of worksheet objects that are part of the workbook.    

**Note:** The Excel REST API supports only Office Open XML file formatted workbooks (files with the`.xlsx` extension). The `.xls` extension workbooks are not supported. 

Microsoft Graph also provides access to Excel capabilities such as [workbook functions](use-functions-in-excel-with-microsoft-graph.md), [updating the display of table ranges](update-range-format-in-excel-with-microsoft-graph.md), and [creating, updating, and displaying charts](display-a-chart-image-in-excel-with-microsoft-graph.md).

## Topics in this section

* [Manage sessions in Excel with Microsoft Graph](manage-sessions-in-excel-with-microsoft-graph.md) 
* [Write to an Excel workbook using Microsoft Graph](write-to-excel-workbook-with-microsoft-graph.md)
* [Use workbook functions in Excel with Microsoft Graph](use-functions-in-excel-with-microsoft-graph.md)
* [Update a range’s format in Excel with Microsoft Graph](update-range-format-in-excel-with-microsoft-graph.md)
* [Display a chart image in Excel with Microsoft Graph](display-a-chart-image-in-excel-with-microsoft-graph.md)
