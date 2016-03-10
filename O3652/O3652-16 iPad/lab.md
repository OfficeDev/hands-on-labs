# Create, test and debug your add-in on a Mac or iPad
In this lab you will learn how to create a sample add-in on the Mac and then side-load it and to test/validate using Excel on the Mac. You will also get hands on experience debugging your add-in on the Mac using Vorlon.js.

**Prerequisites**
1. You must have node.js and VS Code [or your favorite code editor] installed on your Mac.
2. You must also have the latest versions of Office for Mac installed.

## Exercise 1: Create an add-in using Yeoman generator
*In this exercise, you will create a sample Hello World add-in using the Yeoman generator directly on your Mac.*

1. First, install Node.js and Git, if you haven’t already. Next, open a command prompt/terminal as an administrator and run the following command: **npm install –g tsd gulp bower yo generator-office**
2. Create a folder for your add-in project and go to that folder in the command prompt/terminal. Next, run the Yeoman generator for Office Add-ins to create the project scaffolding. Use the following command: **yo office**

When prompted, supply the following information:
- Name of the add-in — myHelloWorldAddin
- Root folder of the project — The current folder
- Type of add-in — Task Pane
- Technology to use — HTML, CSS, and JavaScript
- Supported Office application — Excel (and/or others as desired)

The Yeoman generator creates the structure and basic files for your add-in.

3. Optionally, you can edit the code using your favorite code editor. We recommend VS Code, which includes IntelliSense support when you run the tsd install command from your project folder.

4. Host your add-in. You can host your add-in locally, or use any web server or hosting technology – just make sure that the add-in is served using HTTPS, and update the add-in’s source location in the manifest. To host the add-in using gulp-webserver, use the following command: **gulp serve-static**
You will need to add the self-signed security certificate that gulp-webserver creates as a trusted root certificate or your add-in will not display. To verify that the add-in is running, open your browser and go to the main page at https://localhost:8443/app/home/home.html. 

5. Load the add-in into Office. The easiest way to do this is by sideloading the add-in in Office Online:
- Go to Excel Online and create a blank workbook.
- Go to Insert > Office Add-ins.
- Choose Manage My Add-ins in the upper-right corner of the dialog box, and select Upload My Add-in.
- Select the manifest-myHelloWorldAddin.xml file from the root of the sample project folder, and choose OK. Your add-in will load in Excel Online.





