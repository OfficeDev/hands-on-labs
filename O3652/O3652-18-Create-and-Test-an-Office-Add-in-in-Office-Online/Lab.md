# Create and Test an Office Add-in in Excel Online
This lab will teach you how to create an add-in using only the shell and a browser. 

## Exercise 1: Create an add-in with the Yeoman generator

1. Open a terminal/command prompt.
1. Go to the C:\Dev\ folder on Windows or the ~/Desktop/add-ins directory on Mac.
1. Create a new folder for your add-in project with the command `mkdir <your-name>` and go to that folder.
1. Run the Office yeoman generator by entering the command `yo office`.
1. Provide the following information about your add-in:
  * Name of the add-in: myHelloWorldAddin
  * Root folder of the project: the current folder (press Enter)
  * Type of add-in: Task pane
  * Technology to use: HTML, CSS, & JavaScript
  * Supported Office application: uncheck all options except Excel
1. The yeoman generator will then create all the necessary files for your Excel task pane add-in. You may see some warnings about deprecated components, which you can ignore. When it's done, the add-in can already be used. Run the following command to host the add-in locally: `gulp serve-static`
1. Open a browser and make sure the add-in is working by going to **localhost:8443/app/home/home.html**  (Note: make sure it you navigate to **https://**).
1. If you see a security certificate warning, use Chrome and trust the certificate.

## Exercise 2: Load the add-in in Excel Online

1. Go to office.com and click on the **Excel** tile.
2. Sign in with your Microsoft Account if prompted.
3. Create a blank workbook.
4. Go to the **Insert** tab and choose **Office Add-ins**.
5. In the Office Add-ins dialog, choose **Manage My Add-ins** in the upper-right corner, and select **Upload My Add-in** (if you are logged in to a work or school account, you will see a direct link for "Upload My Add-in" instead of "Manage My Add-ins").
6. Click **Browse** and select the *manifest-myhelloworldaddin.xml* file from your project folder, then click **Upload**.
7. Your add-in should load in Excel. You can type in some data in the spreadsheet, highlight it, and then click the **Get Selected Data** button to see an example that shows how add-ins interact with Office content using the JavaScript APIs.

## Exercise 3: Edit and debug the add-in code

An Office Add-in is just a web app that is displayed within the Office UI and can interact with Office content using Office.js APIs. In this exercise, you'll edit the HTML and JavaScript of the add-in, see your changes reflected, and use a debugger to verify that your code is running properly.

1. Open NotePad (Note: you can use any text editor, we recommend **Visual Studio Code**)
2. Open the home.html file found in your-project-folder/app/home/. 
3. Add a new button after the "Get data from selection" button:

 ```
 <br />
 <button id="write-to-console">Write to debug console</button>
 ```
4. Save the home.html file.
5. Open the home.js file from the same folder.
6. Add a click handler for your new button in the Office.initialize function:

 ```javascript
 Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      jQuery('#get-data-from-selection').click(getDataFromSelection);
      //Add this line:
      jQuery('#write-to-console').click(writeToConsole);
    });
  };
 ```
7. Add the function for the click to perform, which in this case is to write a message to the debug console:

 ```javascript
 function writeToConsole(){
   console.log("Office add-ins are awesome!");
 }
 ```
8. Save the home.js file.
9. Go back to Excel Online and refresh the page. You should see the new button in your add-in.
10. Open the browser's developer tools (this can be done by pressing F12 for most browsers), and go to the Console. Click the button that says "Write to debug console" and you should see "Office add-ins are awesome!" appear in the console.
 

You've now completed the entire lifecycle of add-in development: new project creation, code editing, hosting, loading the add-in into Office, testing, and debugging. You can use this method to create add-ins for any Office application, on any platform that supports add-ins.
