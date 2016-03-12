# Create your first PowerPoint add-in

This quick start will start with an excertize that will guide you through creating a PowerPoint add-in starting from Visual Studio.  
We'll then proceed to excersize 2, where you can start from a pre-existing PowerPoint add-in solution that uses the Excel REST API to connect to a user's OneDrive in Office 365 and retrieve a set of charts from files.

This add-in will expose you to the tools we use to create add-ins and the Office.js 1.1 API. In particular, we'll use the document.setSelectedDataAsync API to insert pictures into the user's current PowerPoint deck.

Let's get going.


## Excersize 1: Create a hello world add-in 

### First, let's create the project.
1. Launch Visual Studio 2015 as administrator.
2. From the **File** menu select the **New Project** command. When the **New Project** dialog appears, select the **PowerPoint Add-in** project template from the **Office/SharePoint** template folder as shown below. Name the new project **HelloWorld** and click **OK** to create the new project.
3. When you create a new **PowerPoint Add-in** project, Visual Studio prompts you with the Choose the add-in type page of the Create Office Add-in dialog. This is the point where you select the type of Add-in you want to create. Leave the default setting with the radio button titled **Add new functionalities to PowerPoint** and select Finish to continue.
4. Visual Studio will create the project. There are a few parts that great created for you:
	- A manifest xml file - this holds the metadata that your add-in needs to run in Office.
	- A web site - The HelloWorldWeb project in the solution contains the html and javascript you need to run your office add-in.
5. When this add in runs it will add a button to the PowerPoint ribbon called "Show Taskpane". 

### Let's change the name of that button to "Hello World"
6. Use the Solution Explorer to drill down into the HelloWorld.xml file as shown in the image below.
 

7. Now, find the xml block that looks like this:
	```XML
	<!-- PrimaryCommandSurface==Main Office Ribbon. -->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!-- Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab. -->
            <OfficeTab id="TabHome">
              <!-- Ensure you provide a unique id for the group. Recommendation for any IDs is to namespace using your company name. -->
              <Group id="Contoso.Group1">
                <!-- Label for your group. resid must point to a ShortString resource. -->
                <Label resid="Contoso.Group1Label" />
                <!-- Icons. Required sizes 16,32,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX. -->
                <!-- Use PNG icons and remember that all URLs on the resources section must use HTTPS. -->
                <Icon>
                  <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                  <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                  <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                </Icon>

                <!-- Control. It can be of type "Button" or "Menu". -->
                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <!-- ToolTip title. resid must point to a ShortString resource. -->
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <!-- ToolTip description. resid must point to a LongString resource. -->
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.tpicon_16x16" />
                    <bt:Image size="32" resid="Contoso.tpicon_32x32" />
                    <bt:Image size="80" resid="Contoso.tpicon_80x80" />
                  </Icon>

                  <!-- This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane. -->
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <!-- Provide a url resource id for the location that will be displayed on the task pane. -->
                    <SourceLocation resid="Contoso.Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
	```
9. Let's modify the button to say "Hello world" instead of "Show Taskpane". Find the following element in the file

	```XML
		<Title resid="Contoso.TaskpaneButton.Label" />
	```
10. This says the label of the title is stored in a string resource named **Contoso.TaskpaneButton.Label**.
11. Scroll down until you find the short string resource with that label.
12. Now, set the DefaultValue attribute to "Hello World".
13. Press F5 to start the project. When Powerpoint loads, you will see a button labeled "Hello World".

Great, let's keep going. In the next section we will learn about how add-ins can interact with the user's Slides.


## Excersize 2: Create an PowerPoint add-in that inserts images of charts from Excel workbooks

For this excersize, we're going to use a pre-existing sample file.
1. Download the following zip file (./quickstart/powerpoint-excel-chart-picker.zip)

2. Unzip that file and open the solution in Visual Studio.  Now, this solution has a bunch of extra code in it that calls the Excel Rest API through the Microsoft Graph.  We won't worry about that for now (there's a link to the Excel REST API quick start at the bottom).

3. When the add-in is running click connect to OneDrive. 

4. Choose a school or work account and enter the following credentials:
- Use this user name:    
- password to login:

5. Once you login, select one of the files and you will see the add-in show the charts in that workbook

Great. Now it's time to plug in the code to insert the image.
6. Stop the project and navigate to file: XXX
7. In the function InsertChart(img), we'll the following code
  
  ```js
    function InsertChart(img) {
      
        //add call to insert the chart into the current slide
    
    }
  ```

