# Create, test, and debug your add-in on a Mac or iPad
In this lab, you will learn how to create a sample add-in on the Mac and then sideload it to test/validate using Excel on the Mac. You will also get hands-on experience debugging your add-in on the Mac using Vorlon.js.

**Prerequisites**

1. You must have node.js and VS Code [or your favorite code editor] installed on your Mac.
2. You must also have the latest versions of Office for Mac installed.

## Exercise 1: Create an add-in using Yeoman generator
*In this exercise, you will create a sample Hello World add-in using the Yeoman generator directly on your Mac.*

1. First, install Node.js and Git, if you haven’t already. Next, open a command prompt/terminal as an administrator and run the following command: **npm install –g tsd gulp bower yo generator-office**
2. Create a folder for your add-in project and go to that folder in the command prompt/terminal. Next, run the Yeoman generator for Office Add-ins to create the project scaffolding. Use the following command: **yo office**
3. When prompted, supply the following information:
Name of the add-in — myHelloWorldAddin
Root folder of the project — The current folder
Type of add-in — Task Pane
Technology to use — HTML, CSS, and JavaScript
Supported Office application — Excel (and/or others as desired). The Yeoman generator creates the structure and basic files for your add-in.
4. Optionally, you can edit the code using your favorite code editor. We recommend VS Code, which includes IntelliSense support when you run the tsd install command from your project folder.
5. Host your add-in. You can host your add-in locally, or use any web server or hosting technology – just make sure that the add-in is served using HTTPS, and update the add-in’s source location in the manifest. To host the add-in using gulp-webserver, use the following command: **gulp serve-static**
6. You will need to add the self-signed security certificate that gulp-webserver creates as a trusted root certificate or your add-in will not display. To verify that the add-in is running, open your browser and go to the main page at https://localhost:8443/app/home/home.html. 

## Exercise 2: Sideload an add-in into Excel for Mac
*In this exercise, we'll go through the process of sideloading an add-in on Excel for Mac.*

1. Open Terminal [shortcut on the desktop] and type **cd ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef**.
2. You might have to create the wef folder if it doesn't already exist.
3. Type **open .** to open Finder
4. Copy the manifest-myHelloWorldAddin.xml file from the root of the sample project folder to this folder.
5. Launch Excel [shortcut on the desktop].
6. Go to the Insert tab and click on the My Add-ins drop-down. Click on the HelloWorld Add-in to open it Excel.

## Exercise 3: Debug HelloWorld add-in using VorlonJS
*In this exercise, we'll use Vorlon.js to debug your add-in on the Mac.*

1. Let's first install by typing **sudo npm I –g vorlon** in Terminal.
2. Type **sudo vorlon** in Terminal to start the VorlonJS server.
3. To verify that Vorlon server is up and running, type http://localhost:1337 in a browser. You should see the VorlonJS start page.
4. Copy the **<script .... ></script>** tag on that page and paste it into the <header> tag. 
5. Launch Excel and start your add-in.
6. You should now see the client connection on the Vorlon server page at http://localhost:1337.





