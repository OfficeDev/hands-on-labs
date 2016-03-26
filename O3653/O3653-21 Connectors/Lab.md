# Learn how to build Office 365 Connectors using Webhooks
Office 365 Connectors are a great way to get useful information and content into your Office 365 Group. From tracking a team's progress in Trello, to following important hashtags in Twitter, Office 365 Connectors make it easier for an Office 365 group to stay in sync with the content they care about and get more done. Developers can build connectors through incoming webhooks to generate rich connector cards.   Connector cards can be short text-based messages, or use "sections" to display rich or specially-formatted information. Outlook takes care of all the UX for you and renders the message automatically. When more content is added to the payload, the card scales gracefully. 

In this lab, you will learn how to post messages to Office 365 Groups using incoming webhooks.  To illustrate using a real world scenario, you will also build an application that receives Github notifications from your favorite repos and post them as connector messages to your Office 365 group. 

## Get an Office 365 developer environment
To complete the exercises below, you will require an Office 365 developer environment. Navigate to https://tryoffice.azurewebsites.net and use the code `BuildChallenge` to get an administrator username and password to one. 

**Pre-requisites**
------------------
 
1. For Exercise 1, you will post a message to the Office 365 Group using simple JSON payload and www.hurl.it tool do the http POST. 

**Environment Prep**
---------

If you don’t have an Office 365 Group, this is how you create one:

1. Sign in to Outlook Web App @ http://Outlook.Office.com using your Office 365 administrator or user credentials.

2. Browse to the Groups and click the **plus (+)** button to create a new Office 365 group. Give a name for the group, and you may keep it public.

3. Click **Connectors** on the menu bar. See the list of available connectors that you can set up for this group.
	



**Exercise 1: Post a Connector Card message (JSON) to the Group**
-------------------------------------------------------------

1. Go to your Office 365 Group. Click **Connectors** on the menu bar, find and expand the Incoming web hook configuration. 
	
2. Click **Add** to create a new configuration, provide a name and click **Create**. This will generate a webhook URL for the group.  
	
3. Copy the generated URL and save it-as you will need it later. Select **Done** to create the incoming webhook.
	
4. Open a new browser tab and navigate to https://www.hurl.it, which is an in-browser web request composer similar to what Fiddler offers.

5. When the page loads, add the following details:
	- **Operation**: **POST**
	- **Destination Address**: **paste the webhook URL** from **Step 3**
	- **Headers**: **Content-Type: application/json** 
	- **Body**: **{ "text": "Hello from Build 2016" }**
	- **Copy and paste the Sample Connector Card Message JSON payload (see below these instructions)**

	![Manual Webhook](http://i.imgur.com/vV8FKeD.png)
	

6. Accept the Captcha and click the **Launch Request** button. You should get a confirmation screen that looks similar to the following.

	![Webhook Manual Confirmation](http://i.imgur.com/LjEi7m6.png)


7. Go back to your Office 365 Group. Review the connector card message in the Group inbox that was sent using the incoming webhook. 

8. Optional: Change the JSON card format to further customize the layout, buttons, and colors.  For example, under potential actions, rename the button or change the target URL.

9. To learn more about the connector card format, visit https://dev.outlook.com/Connectors/GetStarted.
	
	
### Sample Connector Card Message JSON Payload

	{  
	  "summary": "New Comment by Ben Quillen on \"Fabrikam Forum\"",
	  "title": null,
	  "text": null,
	  "themeColor": "#3479BF",
	  "sections": [
	    {
	      "title": null,
	      "text": null,
	      "markdown": true,
	      "facts": [
	        {
	          "name": "Added By",
	          "value": "Ben Quillen"
	        },
	        {
	          "name": "Date",
	          "value": "15-12-2015"
	        },
	        {
	          "name": "Priority",
	          "value": "Medium"
	        },
	        {
	          "name": "State",
	          "value": "Active"
	        }
	      ],
	      "images": null,
	      "activityTitle": "Ben Quillen commented",
	      "activitySubtitle": "on \"Fabrikam Forum\"",
	      "activityText": "We should prioritize this effort.",
	      "activityImage": "https://cdn0.iconfinder.com/data/icons/PRACTIKA/256/user.png"
	    }
	  ],
	  "potentialAction": [
	    {
	      "@context": "http://schema.org",
	      "@type": "ViewAction",
	      "name": "View details",
	      "target": [
	        "http://microsoft.com"
	      ]
	    }
	  ]
	}
	

**Exercise 2: Build an ASP.net application to receive incoming notifications from Github service and post them as connector card messages in Office 365 groups**
------------------------------------------------------------------------
	
1. Start Visual Studio 2015 and select **New>Project**. Under Templates for Visual C#, select **Web** and select the **ASP.net Web application** template. Provide a project name. Select **empty** for the ASP.net template type and select the **Web API** checkbox. On the right, look under Microsoft Azure and uncheck the **Host in the Cloud** checkbox.  This will pull down the necessary nugets for you.

2. Click on **Tools>Extensions and Updates** in Visual Studio, select Online on left and in the search box, search for ASP.NET Webhooks Connected Service. If you do not have ASP.Net WebhHooks Connected Services installed, follow the steps to download and install it. Now right-click on the Visual Studio project, select **Add>Connected Service**. Select ASP.Net on the left, select **ASP.net WebHooks** click configure. OThis will generate code into your project, so you can host and receive webhooks.

3. Select Github in the **Enable incoming webhooks** dialog box. Provide a secret for your GitHub application. Let's generate a new secret. Go to sha1-online.com and generate a new SHA key. Copy this key into your clipboard and also save it in Notepad temporarily. Paste this secret into Github text box in Visual Studio.

4. Visual studio will install the packages for Github webhook receivers and generate all the necessary code in your project. Verify the GitHub secret is present in the Web.config file of your project by looking under app settings for the key listed as `"MS_WebHookReceiverSecret_GitHub".` 

5. Open the global.asax file. Initialize the webhook receiver. Add a line of code for `GlobalConifguration.Configure(WebhookConfig.Register)`  in the Application_Start function.

		protected void Application_Start()
	        {
	            GlobalConfiguration.Configure(WebApiConfig.Register);
	            GlobalConfiguration.Configure(WebHookConfig.Register);
	        }

6. Navigate to “C:\git\trainingcontent-nda\O3653\O3653-21 Connectors\Completed Projects\Github Webhooks\WebApplication1\Models\” in Windows Explorer. 

7. Copy all the files in this directory that begin with “Swift” to the Models directory in your Visual Studio project. Do the same for  GithubIssueEvent.cs, ConnectorCard.cs, ModelBuilder.cs and copy them to the same Models directory. In Visual Studio, right click on the Models folder, select **Add>Existing** item and add all these files into your Visual Studio project.
	        
8. Open the GitHubWebHookHandler.cs file under WebHandlers folder in your Visual Studio project.  Replace the existing code in this file with the code listed below. Fix the namespace (see curly braces), so it continues to bear the same name as your Visual Studio project.

		using Microsoft.AspNet.WebHooks;
		using Newtonsoft.Json.Linq;
		using System.Linq;
		using System.Threading.Tasks;
		using Newtonsoft.Json;
		using GithubTest;
		using System.Net.Http;
		using System;
		using System.Text;
		
		namespace MyTestConnectorForGithub.WebHookHandlers
		{
		    public class GitHubWebHookHandler : WebHookHandler
		    {

		        // TODO: Copy and paste the group webhook URL here.
		        public const string groupWebHookURL = @"paste the Office 365 Group webhook URL here";

		        public override Task ExecuteAsync(string receiver, WebHookHandlerContext context)
		        {
		        // Be sure we're only processing the intended type of hook.
		        if("GitHub".Equals(receiver, System.StringComparison.CurrentCultureIgnoreCase))
		        {
		                // TODO: Replace this placeholder functionality with your own code.
		                string action = context.Actions.First();
		                JObject incoming = context.GetDataOrDefault<JObject>();
		                string connectorCardPayload = ConnectorCard.ConvertGithubJsonToConnectorCard(incoming.ToString());
		                var body = PostRequest(connectorCardPayload).Result;
		            }
		
		            return Task.FromResult(true);
		        }
		
		
		        private static async Task<HttpResponseMessage> PostRequest(string payload)
		        {
		            var targetUri = new Uri(groupWebHookURL);
		            var httpClient = new HttpClient();
		
		            return await httpClient.PostAsync(targetUri,
		                 new StringContent(payload, Encoding.UTF8, "application/json"));
		        }
		
		    }
		}
	
	
9. In the same GitHubWebHookHandler.cs file, find the string variable groupWebHookURL.  Copy and paste the Office 365 group webhook URL that you previously got from Exercise #1 (Step 3).

10. Build your project. Get ready to publish your project
 
	**Publish the app using trial Azurewebsites service**

	**Note:** If you already have an Azure subscription and familiar with publishing your website to Azurewebsites, then right-click on the Visual studio project and select **Publish**. Select the Microsoft Azure Web Apps option and sign in to your Azure subscription (as  needed) and follow simple steps to host ASP.net application in Azurewebsites. 
	
	Browse to https://tryappservice.azure.com/ to create a temporary test site. Choose **Web App** as the app type, then choose **Next**. Change the language dropdown to **C#**, then choose **ASP.NET Empty Site** and choose **Create**. Sign in with an account 	to complete the creation process.
	
	When the site is created, copy the site URL, then choose the **Download publishing profile** link and save the file to the local machine. 
	
	In Visual Studio, open the **Build** menu and choose **Publish**.
		
	Select **Import** under **Select a publish target**. Browse to the publishing profile you downloaded in the previous step. Choose the **Validate Connection** button to make sure the settings work.
	
	Choose the **Settings** item in the left navigation. Select **Debug**. Expand **File Publish Options** and put a check in 	the **Remove additional files at destination** checkbox.
		
	Choose **Publish** to publish the app to Azure. Once the publishing process is complete, a new browser window will open to the newly published site.

11. Go to github.com. Create a new github repo (unless you already have one for testing purposes). Navigate to the repo. Click  **Settings**. Select Webhooks & Services on the left side. Click the **Add webhook** button. Provide the full webhook URL of your ASP.net application. To get this url, append "/api/webhooks/incoming/Github" to the URL of your Azure web application (as an example, the full webhook URL would be http://mywebhookspreview.azurewebsites.net/api/webhooks/incoming/Github). You will receive notifications from Github service at this URL.  

12. For the secret, enter the SHA1 key you got earlier.

13. Under "Which events would you like to trigger this webhook", select "Let me select individual events". Select the "Issues" checkbox. You will receive notification only when issues are created, assigned, labelled, or closed. For this exercise, we will focus on the experience for issue creation. Click the **Update webhook** button to save your Github webhook configuration.

14. Optional: To debug the incoming webhook from Github service, place a breakpoint on the ExecuteAsync function of the GithubWebhookHandler.cs file of your Visual Studio project. To do this, open server explorer in Visual Studio, find and right-click your web application under App service, and attach the debugger. 

15. Create a new issue for your repo. This will trigger the incoming webhook in your ASP.net application. You should now receive connector card messages for new github issues in your Office 365 group inbox.

**Congratulations on completing this exercise!** Try building your own connector and submit it to Office 365. Customize the card message experience to show your own brand and style, and the sender avatar will show your logo (instead of "incoming webhook"). 
Check out dev.outlook.com/connectors for more developer information, code samples, and instructions for submitting and listing your connector in the Office 365 Connectors catalog.


