# Create your first Office add-in with the Outlook JavaScript APIs

In this lab, you will use Visual Studio to create your first Outlook add-in using the Outlook JavaScript APIs. The add-in will allow the user to translate parts of a message they are composing into different languages.

## Prerequisites

1. You must have an Office 365 tenant to
   complete this lab. If you do not have one, the lab for **O3651-7 Setting up
   your Developer environment in Office 365** shows you how to obtain a trial.
1. You must have Visual Studio 2015 with Update 1 installed.
1. You must have the Microsoft Office Developer Tools for Visual Studio 2015 installed.
1. You must have Outlook 2016 installed.

## Exercise 1: Create a new Outlook add-in project

In this exercise you will create a new project using the Outlook add-in template.

1. Launch Visual Studio 2015 and select **New**, **Project**.
  1. Expand **Templates**, **Visual C#**, **Office/SharePoint** ,**Web add-ins**. Select **Outlook Add-in**. Name the project **Translator** and click **OK**.
  
    ![The new project dialog using the Outlook add-in template](./Images/create-project.PNG)
  
1. Run the app to verify it works.
  1. Press F5 to begin debugging.
  1. When prompted, enter the email address and password of your Office 365 account. Visual Studio will install the add-in for that user.
  
    ![The Connect to Exchange email account dialog](./Images/deploy-addin.PNG)
    
  1. With the app running, open Outlook 2016 and logon to the user's mailbox. You should see a **Display all propeties** button on the ribbon when you select or open a message.
  
    ![A message in Outlook 2016 with the add-in button on the ribbon](./Images/default-button.PNG)
    
Now that you've verified that the add-in is working, exercise 1 is complete!

## Exercise 2: Add translation