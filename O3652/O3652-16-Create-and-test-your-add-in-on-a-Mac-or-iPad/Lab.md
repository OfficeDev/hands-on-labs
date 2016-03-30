# Create, test, and debug your add-in on a Mac or iPad
In this lab, you will learn how to create a sample add-in on the Mac and then sideload it to test/validate using Excel on the Mac. You will also get hands-on experience debugging your add-in on the Mac using Vorlon.js.

**Prerequisites**
The following prerequisies have already been installed on Mac you are currently working on:
1. You must have node.js and VS Code [or your favorite code editor] installed on your Mac.
2. You must also have the latest versions of Office for Mac installed.

## Exercise 1: Create an add-in using Yeoman generator
*In this exercise, you will create a sample Hello World add-in using the Yeoman generator directly on your Mac.*

1. Create a folder for your add-in project in the ~/Desktop/add-ins folder and go to that folder in the command prompt/terminal. 
2. Run the Yeoman generator for Office Add-ins to create the project scaffolding. Use the following command: `yo office`
3. When prompted, supply the following information:
  * Name of the add-in — **yournameHelloWorldAddin**
  * Type of add-in — **Task Pane**
  * Technology to use — **HTML, CSS, and JavaScript**
  * Supported Office application — **Excel** (and/or others as desired). The Yeoman generator creates the structure and basic files for your add-in.
4. Optionally, you can edit the code using your favorite code editor. We recommend VS Code, which includes IntelliSense support when you run the tsd install command from your project folder.
5. Host your add-in using gulp-webserver by using the following command: `gulp serve-static`
6. To verify that the add-in is running, open your browser and go to the main page at [https://localhost:8443/app/home/home.html](https://localhost:8443/app/home/home.html)

## Exercise 2: Sideload an add-in into Excel for Mac
*In this exercise, we'll go through the process of sideloading an add-in on Excel for Mac.*

1. Open Terminal [shortcut on the desktop] and type

    `cd ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`

2. You might have to create the wef folder if it doesn't already exist.
3. Type `open .` to open Finder
4. Copy the `manifest-yournameHelloWorldAddin.xml` file from the root of the sample project folder to this folder.
5. Launch Excel [shortcut on the desktop].
6. Go to the Insert tab and click on the My Add-ins drop-down. Click on the HelloWorld Add-in to open it Excel.

## Exercise 3: Debug HelloWorld add-in using VorlonJS
*In this exercise, we'll use Vorlon.js to debug your add-in on the Mac.*

1. Go to ~/Desktop/add-ins folder
2. Type `sudo vorlon` in Terminal to start the VorlonJS server.
3. To verify that Vorlon server is up and running, type `https://localhost:1337` in a browser. You should see the VorlonJS start page.
4. Copy the `<script .... ></script>` tag on that page and paste it into the `<head>` tag of `app/home/home.html`. 
5. Launch Excel and start your add-in.
6. You should now see the client connection on the Vorlon server page at `https://localhost:1337`.
7. Click on the client connect link to view the Vorlon debugger UI. You can now use the Dom Explorer and Obj. Explorer tabs to view/edit the source code on your add-in.

**Optional**
Install the office.js plug-in for Vorlon and use it to test Office APIs from this blog post: http://blogs.msdn.com/b/mim/archive/2016/02/18/vorlonjs-plugin-for-debugging-office-addin.aspx

## Exercise 4: Edit the add-in code

An Office Add-in is just a web app that is displayed within the Office UI and can interact with Office content using Office.js APIs. In this exercise, you'll edit the HTML and JavaScript of the add-in to get a sense for the entire lifecycle of an add-in project.

1. Launch VS Code.
2. Open the home.html file found in ~/Desktop/add-ins/your-project-folder/app/home/. 
3. Add a new button after the "Get data from selection" button:

 ```
 <br />
 <button id="write-data-to-selection">Write data to selection</button>
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
      jQuery('#write-data-to-selection').click(writeDataToSelection);
    });
  };
 ```
7. Add the function for the click to perform, which in this case is to write a message to the current location in the document:

 ```javascript
 function writeDataToSelection(){
     Office.context.document.setSelectedDataAsync("Office add-ins are awesome!",
      function(result){
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          app.showNotification('Data successfully written.', "");
          console.log("Writing to the document succeeded!");
        } else {
          app.showNotification('Error:', result.error.message);
          console.log("Writing to the document failed: " + result.error.message);
        }
      }
    );
 }
 ```
8. Save the home.js file. Console.log statements above should be visible in the VorlonJS console.
9. Go back to Excel, click on the "i" in the top right corner of the add-in pane and then select "Reload". You should see the new button in your add-in.
10. Select an empty cell in the worksheet and click the new button that says "Write data to selection". You should see "Office add-ins are awesome!" written to the cell.
 

You have now completed building and debugging a new add-in entirely on the Mac. This Office add-in will run on all platform where Office supports add-ins.

