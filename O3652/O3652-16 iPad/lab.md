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
Name of the add-in — myHelloWorldAddin
Root folder of the project — The current folder
Type of add-in — Task Pane
Technology to use — HTML, CSS, and JavaScript
Supported Office application — Excel (and/or others as desired)
 
The Yeoman generator creates the structure and basic files for your add-in.



