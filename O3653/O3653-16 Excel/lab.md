# Excel REST API 

The purpose of this lab is to show step-by-step instruction for reading writing into an Excel document stored in your OneDrive Business using new Excel REST APIs. We'll use Visual Studio MVC project to showcase the interaction. 

## Usage scenario

* User maintains or stores Excel files in OneDrive Business. 
* User authorizes a web/mobile application to read/update OneDrive file contents.
* App can access the Excel file contents and make updates over the OneDrive files REST API available through Microsoft Graph. 
* A `/workbook` segment is added in the URL at the end of file identifier to distinguish the Excel API call and access workbook's data model. Example: 
https://graph.microsoft.com/v1.0/me/files/{id}/workbook/worksheets/Sheet1/tables
* Using Excel REST API, App reads or updates the Excel file content as necessary. Any updates made to the file is saved to the document on OneDrive. 

Note: 

* A workbook corresponds to one Excel document. Only one document can be addressed at a time. 
* The Excel API doesn’t allow user to create or delete the document itself. For those functionalities, regular OneDrive files API can be used. 


## Platform support

Currently, the Excel REST APIs are supported on any Excel workbook stored on your OneDrive Business document library or Group's document library. 

## Excel REST object model

There are several resources available as part of Excel API. Below list shows some of the top level important objects.

* Workbook: Workbook is the top level object which contains related workbook objects such as worksheets, tables, named items, etc.
* Worksheet: The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
* Range: Range represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells.  
* Table: Represents collection of organized cells designed to make management of the data easy. 
	* TableColumn: Represents a column in the table
	* TableRow: Represents a row in the table. 
* Chart: Represents a chart object in a workbook, which is a visual representation of underlying data.  
* NamedItem: Represents a defined name for a range of cells or a value. Names can be primitive named objects (as seen in the type below), range object, etc.
* Application: Represents the Excel application that manages the workbook. Get the calculation mode of the workbook and perform calculation.
* Create Session: Create Excel workbook sessions. It is a good practice to create workbook session and pass it along with the request as part of the request header as it allows the server to link the API request to an existing in-memory copy of the file on the server. If a session ID is not provided, the server dynamically creates a session behind the scene. However, this requires additional server side processing and could add to the latency of the response. Session ID has a life span which gets extended with each usage or regresh. Once a session ID has expired, a new session session ID needs to be created. If an expired or invalid session token is provided as part of the request, the API will return an error indicating that the session ID is not valid. 		

## Lab instructions 
