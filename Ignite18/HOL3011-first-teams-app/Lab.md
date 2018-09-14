# 0-60: Get up and running with your very first Microsoft Teams app
## Overview and Introduction
In this lab, you will be getting your first app up and running in Teams. You should do this lab if:
You have used Microsoft Teams
You have built add-ins for Office or SharePoint
You are interested in building apps or solutions for Teams, or Office 365 in general
We will start with an app that provides basic bot and tab functionality already hosted on Azure. The lab will then include instructions for adding new functionality through the app’s manifest. For those who want to go further, you can add enhancements to the bot and deploy those changes to your own Azure subscription. 
The sample app we're using here manages the hiring process of candidates for open positions in a team – a Talent Management Application. While it looks good, the app doesn't actually do anything – we want to focus on building a Teams app and loading it into Teams, not creating a full-blown talent management application.
## Prerequisites
Your lab environment should already come preinstalled with the following tools
- Ngrok
- Visual Studio
- Microsoft Teams desktop client
- Browser of your choice
# PART 1: Get familiar with Teams apps
In this first part of the Lab you’re going to create a Teams app using an existing base manifest provided as part of this lab. To get you up and running, we will provide a web app that is already running in Azure.
The app package **teams-sample-app-package.zip** is located in this folder. Download it to your PC and unzip the contents to any local folder. It will contain three files:
- manifest.json
- color_icon.png
- outline_icon.png
## Step 1: Prepare your Office 365 environment
You will first need to enable side loading for Teams within your Office 365 environment. Open the Admin Center by visiting https://admin.microsoft.com/AdminPortal/Home#/Settings/ServicesAndAddIns from your browser.

Next, select Microsoft Teams. Under the Apps section, scroll down to External Apps and make sure that “Allow sideloading of external apps” is set to On.

## Step 2: Create your app using Teams App Studio
The Teams desktop client is pinned to the Task Bar. Click on it and log in with your VM’s credentials. You can play around with Teams and create your own teams and channels.
To create the app package, Teams has a tool called App Studio – and it's actually a Teams app itself. Install it from the Teams app store:

Image

Click on the "Store" icon at the lower left, search for "app studio", click on the "Teams App Studio" entry, "Install" button on the consent dialog, and then the bottom "Open" button on the second next dialog:
Image 

Click on the "Manifest editor" tab and the "Import and existing app" button:
Image

Load the manifest.json file that you previously unzipped and then click on the "Contoso Talent" entry.

Image
 
Most of the information has already been filled out for you. The following screenshots show what information to change:

### App Details
Image

Tabs
A tab is an embedding of an existing web application experience inside of Teams, which users can collaborate around.
You will update the tab information to point to the Azure-hosted tab. In the later section, you can update these entries to point to your local setup.
 
### Bots
This app supports a bot that users can interact with through natural language. It supports a number of commands the return rich cards.
You will update the bot information to point to the Azure-hosted instance. In the later section, you can update these entries to point to your local service (through Ngrok).
 
You can also optionally add a new command. The bot we provide supports signing into AAD and the MS Graph. Under Commands, click “Add” and enter in “login” as the command text, with “connect to Office 365” as the help text. Check both Personal and Team scopes.
 
### Messaging extensions
Messaging extensions allow users to search your backend through the Teams UI, similar to how users can query Giphys and emojis. These results are returned as rich cards that can be posted into channels.
 
## Step 3: Run the Sample App
You can load and test your sample app directly from App Studio. To do this, click “Test and distribute” under the Finish section in the Manifest editor. Click “Install” and select the team in which you want to test the app.
 
Image
 
Next, you'll see the dialog below (of course, the team name will be different). Here, it shows the General Channel:
Image

You're now free to experiment with your app:
- Use the "Personal App" version via the "…" menu on the left side of Teams
- Talk to the bot in both 1:1 and channel mode
- Use actionable messages to schedule interviews
- Create tabs and add them to channels
- Use the messaging extension to find candidate cards to enrich your conversations

Alternatively, you can also upload your custom app through the Teams UI. To do this, first download your app package locally. It will be saved to your PC’s Downloads folder.

 

Next, click on the Store icon in the Teams client and then click "Upload a custom app" at the lower left – the file will be located in your Downloads folder and it's called teams-sample-app-package.zip (if you are using the Azure version) or ContosoTalent.zip if you built it yourself.
 
# PART 2: Deploying and testing locally
In Part 2 of the lab, you’ll get a chance to run the app locally, make some minor changes, and then reload the app in Teams. First, you’ll need to grab the source code to run locally by running this git command:
git clone https://github.com/billbliss/microsoft-teams-sample-talent-acquisition
Open the solution in Visual Studio by double-clicking on the .sln file. Leave it open for now – we’ll come back to it later.
Step 4: Create a bot through Bot Framework (Optional)
Next, you need to register a bot through the Bot Framework portal. Navigate in your browser to https://dev.botframework.com/bots/new 
Click on the "Sign in" button and log on with your demo tenant or MSA credentials. Agree to the Terms and Conditions if necessary, and you should see a page that looks like what's below. Fill it in according to the instructions.
 

Once you've logged in to the Application Registration portal (https://apps.dev.microsoft.com), the App name you just created in Bot Framework will appear and an App ID will be generated. Copy this to the clipboard and paste it into Notepad. Then generate a password and copy/paste that to Notepad too. If you forgot to copy the password, simply generate a new one. Remember to click on the "Finish and go back to Bot Framework" button because you're not done yet!
 
This will take you back to the previous page. Scroll down to the bottom and click the checkbox and click the Register button. That will take you to a page that looks like this:
 
Click on the Microsoft Teams icon to add it as a channel (which in this context, has nothing to do with Microsoft Teams channels). Agree to the Terms of Service and you'll see a "Configure MSTeams" – click on Done at the lower left (there's nothing to configure):
 
After you press the Done button, you'll see Microsoft Teams added to the list of channels. Leave the "Connect to channels" page for your bot open – we're going to come back to it shortly.
Step 5: Set your App ID and Password and test your bot
Return to Visual Studio and open the Web.config file at the root of the solution. In the TeamsAppId/MicrosoftAppId/MicrosoftAppPassword sections, copy/paste the App ID and Password from Notepad. TeamsAppId doesn't have to be the same as MicrosoftAppId, but it's usually easier if it is, so use the same App ID for both. When you are done, it should look like this:
 
Save the Web.config file and run your solution again.
Now, return to the "Connect to channels" page for your bot, and press the "Test" button at the upper right:
 
Type "hello" at the lower right and your bot should respond (if a "retry" link appears next to what you typed, click it):
 
We've verified that your bot is working, so let's try it in Teams.
 

## Step 6: Tunnel localhost to the Internet
Although a Microsoft Teams app is free to access information and APIs inside your firewall, some portions of it, such as the tab URL and bot endpoint, must be accessible from the Internet. The app that you will create today will be running on localhost, so we need a way to make code running on your local machine be accessible from the Internet.
We're using a tool called Ngrok (ngrok.com) for this purpose. 
In the open command prompt (or you can start a new one), type the following command (if ngrok isn't in your PATH, you'll have to prepend its installation directory):
>ngrok http 3979 -host-header=localhost:3979

After a bit, you should see something like this, although the http/https URLs will be different:
 
Copy the https: URL (not the http: URL) to the clipboard. In the example above, it's https://b26d0449.ngrok.io, but of course yours will be different. **Save the URL; you'll need it shortly.** 
You can minimize the ngrok window; it will keep running in the background.
 

## Step 7: Start your app in Visual Studio
Next, we're going to make a quick check that everything is working properly in Visual Studio. Switch to Visual Studio and click on the Run icon:
 
Visual Studio will build the solution and open http://localhost:3979. But we're interested in what's on the Internet, so paste the URL you saved earlier into a new browser tab. You should see the same page:
 
You can stop the app now or leave it running and stop it later.
## Step 8: Update your app package and test
In Part 1 of this lab, you used our Azure instance to test your app. Now, you can use your locally-running Ngrok instance. Return to App Studio within Teams and update these values:
### Tabs
Update the team tab configuration URL and personal tab URL to point to your ngrok instance, e.g. https://b26d0449.ngrok.io/channelconfig.html
 

### Bots
Paste in the bot ID that you registered in Step 5. You can also use App Studio to manage your bot’s endpoint by following the instructions in “Setup”.
 
### Messaging extensions
Similar to bots, update the bot ID to be the one you obtained in Step 5.
   
Now, from App Studio, you can side load the app 
## Step 9: Make some code changes
In this step, you’ll add a new command to the messaging extension to allow searching for candidates in addition to open positions.
Update your code
In Visual Studio’s solution explorer, under the **Messaging** folder, open **MessagingExtension.cs**
Under the CreateResponse() method, add the following block of code at line 93:
```csharp
else if (query.CommandId == "searchCandidates")
{
  string name = query.Parameters[0].Value.ToString();
  CandidatesDataController controller = new CandidatesDataController();

  foreach (Candidate c in controller.GetTopCandidates("ABCD1234"))
  {
    c.Name = c.Name.Split(' ')[0] + " " + CultureInfo.CurrentCulture.TextInfo.ToTitleCase(name);
    var card = CardHelper.CreateSummaryCardForCandidate(c);

    var composeExtensionAttachment = card.ToAttachment().ToComposeExtensionAttachment(CardHelper.CreatePreviewCardForCandidate(c).ToAttachment());
    results.Attachments.Add(composeExtensionAttachment);
  }
}
```
This block of code is what responds to the new command to search for candidates. Rebuild your solution and rerun by hitting F5. In the next step you’ll wire up the command to your app’s manifest.
### Add a new command
Now you’ll add a new command under the Messaging extensions section of your app in App Studio. Provide the following field values:
- Command Id = searchCandidates
- Title = Candidates
- Description = \<whatever string you want\>
- Parameter
- Name = name
- Title = Name
- Description = \<whatever string you want\>
 
## Test your app
In Visual Studio, hit F5 to restart your local service.

Under Test and distribute, click Install to reload your app.

In Teams, go to any chat or channel conversation. Click on the “…” below the compose box to open the Contoso Talent app – you should now see your new command. Type in any string to initiate the search with your new code changes.
