# Create your first PowerPoint add-in

This quick start will guide you through creating a PowerPoint add-in starting from Visual Studio.  
We'll then start from a pre-existing add-in solution that uses the Excel REST API to connect to a user's OneDrive in Office 365 and retrieve a set of charts from files.

This add-in will expose you to the tools we use to create add-ins and the Office.js 1.1 API. In particular, we'll use the document.setSelectedDataAsync API to insert pictures into the user's current PowerPoint deck.
To start, you will need Visual Studio


## Excersize 1: Create a simple hello world add-in 
1. Open Visual Studio 2015
2. File > New > Project > Office > Office Add-in
3. Choose Taskpane add-in and choose PowerPoint only as the add-in (this will make it easy for Visual Studio to detect that it should launch PowerPoint when pressing F5).
4. The template will create a Manifest file that declares
5. In addition it will

Go ahead and press F5. When PowerPoint opens, click the Insert tab and select the image.

## Excersize 2: Create the Excel Chart picker add-in for PowerPoint

For this excersize, we're going to use a pre-existing sample file.

1. Download the following zip file (./quickstart/powerpoint-excel-chart-picker.zip)
2. Unzip that file and open the solution in Visual Studio.  Now, this solution has a bunch of extra code in it that calls the Excel Rest API through the Microsoft Graph.  We won't worry about that for now (there's a link to the Excel REST API quick start at the bottom).
3. Press F5 and run the template. 
4. When the add-in is running click connect to OneDrive. 
5. Choose a school or work account and enter the following credentials:
- Use this user name:    
- password to login:
6. Once you login, select one of the files and you will see the add-in show the charts in that workbook

Great. Now it's time to plug in the code to insert the image.
7. Stop the project and navigate to file: XXX
8. In the function InsertChart(img), we'll the following code
  
  ```js
    function InsertChart(img) {
      
        //add call to insert the chart into the current slide
    
    }
  ```

