# Learn how to build Office 365 Connectors using Webhooks
Office 365 Connectors are a great way to get useful information and content into your Office 365 Group. Any user can connect their group to services like Trello, Bing News, Twitter, etc., and get notified of the group's activity in that service. From tracking a team's progress in Trello, to following important hashtags in Twitter, Office 365 Connectors make it easier for an Office 365 group to stay in sync and get more done. Developers can build connectors through incoming webhooks to generate rich connector cards.   Connector cards can be short text-based messages, or use "sections" to display rich or specially-formatted information. Outlook takes care of all the UX for you and renders the message automatically. When more content is added to the payload, the card scales gracefully. 

In this lab, you will use tools like Postman (or Fiddler) to post messages to Office 365 Groups using incoming webhooks. 
To illustrate using a real world scenario, you will also build an application that receives Github notifications from your favorite repos and post them as connector messages to your Office 365 group. 


**Pre-requisites**
------------------

 1. You must have an Office 365 tenant and Microsoft Azure subscription to complete this lab. If you do not have one, the lab for O3651-7 Setting up your Developer environment in Office 365 shows you how to obtain a trial.
 
 2. For Exercise 1, use tools like Fiddler, Postman or curl to post a JSON payload to the group's webhook URL. This exercise will use Postman to post JSON message. You may also use the Connector playground sandbox  to complete this exercise
 
 3. For Exercise 2, you must have Visual Studio 2015 with Update 1 installed. 


**Environment Prep**
---------

If you don’t have an Office 365 Group, this is how you create one

1. Sign in to Outlook Web App @ http://Outlook.Office.com using your Office 365 administrator or user credentials

2. Browse to the Groups and click the "+" button to create a new Office 365 group. Give a name for the group, and you may keep it public

3. Click Connectors on the menu bar. See the list of available connectors that you can set up for this group.
	

**

**Exercise 1: Post a Connector Card message (JSON) to the Group**
-------------------------------------------------------------

1. Go to your office 365 Group. Click Connectors on the menu bar, find and expand the Incoming web hook configuration. 
	
2. Click Add to create a new configuration, provide a name and click create. This will generate a webhook URL for the group.  
	
3. Copy the generated URL and save it someplace as you will need it later.  Select Done to create the incoming webhook.
	
4. Launch the Postman application, copy and paste the incoming webhook URL (from the previous step) into the POST text field. Select body (raw) and application/json for payload
	
5. Copy and paste the sample payload below into the body of Postman and select the Send button
 
6. Review the connector card message in the Group inbox that was sent using the incoming webhook 

7. Optional: change the JSON card format to further customize the layout, buttons & colors.  For e.g. under potential actions, rename the button or change the target URL

8. To learn more about the connector card format, visit https://dev.outlook.com/Connectors/GetStarted
	
	
Sample Connector Card Message JSON Payload

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
	
1. Launch Visual Studio 2015 and select New, Project. Select the ASP.net Web application template.  Provide a name for your project. Select empty for the ASP.net type and select the Web API checkbox. This will automatically pull down the necessary nugets for you.

2. Right click on the project, select "Add" -> "Connected Service". Select ASP.net on the left side of the dialog and select the ASP.net Webhooks. Click the configure button. This will generate code into your project so you can host and receive webhooks.

3. Select Github in the "Enable incoming webhooks" dialog. You need to provide a secret for your GitHub application. Let's generate a new secret. Go to sha1-online.com and generate a new SHA key. Copy this key into your clipboard and also save it in notepad temporarily. Paste this secret into Github text box in Visual Studio.

4. Visual studio will install the packages for Github webhook receivers and generate all the necessary code in your project. Verify the GitHub secret is present in the Web.config file of your project, by looking under app settings for the key listed as `"MS_WebHookReceiverSecret_GitHub".` 

5. Open the global.asax file. You need to initialize the webhook receiver. Add a line of code for `GlobalConifguration.Configure(WebhookConfig.Register)`  in the Application_Start function. 

		protected void Application_Start()
	        {
	            GlobalConfiguration.Configure(WebApiConfig.Register);
	            GlobalConfiguration.Configure(WebHookConfig.Register);
	        }
	        
6. Find the "Completed Projects" folder under lab "O3653-21". You will find this folder in the same Github repo location that host the instructions for this exercise.  You will need to copy a few files from this already completed projects folder in order to complete this exercise.

7. Find all the Swift*.cs files under the Models folder in "Completed Projects". Copy these files to the Models folder in your Visual Studio project.  Right click on the Models folder, select Add -> "Existing item" and add these files into your Visual studio project.  Do the same for GithubIssueEvent.cs, ConnectorCard.cs and add them to your visual studio project. 

8. Open the GitHubWebHookHandler.cs file under WebHandlers folder in your visual studio project.  Replace the existing code in this file with the code listed here below. Fix the namespace (see curly braces), so it continues to bear the same name as your Visual studio project

		using Microsoft.AspNet.WebHooks;
		using Newtonsoft.Json.Linq;
		using System.Linq;
		using System.Threading.Tasks;
		using Newtonsoft.Json;
		using GithubTest;
		using System.Net.Http;
		using System;
		using System.Text;
		
		namespace {Use the project name as namespace}.WebHookHandlers
		{
		    public class GitHubWebHookHandler : WebHookHandler
		    {

		        // TODO: Copy and paste the group webhook URL here
		        public const string groupWebHookURL = @"paste the Office 365 Group webhook URL here";

		        public override Task ExecuteAsync(string receiver, WebHookHandlerContext context)
		        {
		        // make sure we're only processing the intended type of hook
		        if("GitHub".Equals(receiver, System.StringComparison.CurrentCultureIgnoreCase))
		        {
		                // todo: replace this placeholder functionality with your own code
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
	
	
9. In the same GitHubWebHookHandler.cs file, find the string variable groupWebHookURL.  Copy and paste the Office 365 group webhook URL that you previously got from Exercise #1. 

10. Build your project.  Right click on the project and select "Publish".  
	
	a. Select the Microsoft Azure Web Apps option and sign in to your Azure subscription (if needed). 
	
	b. Create a new web app.  Choose a web app name and app service plan location. Click Create.  Select  "Settings" on the left side, choose the Debug configuration (so you can debug your web application in Visual Studio) and select all the checkboxes under file publish options. 
	
	c. Click Publish to publish the webapp to Azure websites. This will launch a browser and take you to the Azure websites hosting your web application (e.g. http://mywebhookspreview.azurewebsites.net) Copy this URL.
	
		Note: If you don’t have an Azure subscription, you can get it free by signing up @ https://tryappservice.azure.com/ This allows you to host you web application on Azure for up to 24 hours, no credit card  required.   Choose **Web App** as the app type, then click **Next**.  Change the language dropdown to **C#**, then choose **ASP.NET Empty Site** and click **Create**.  Choose to download publishing profile.  When using Visual Studio to publish your project, choose the import option to import the publishing profile, follow the same a, b and c steps above
	
11. Go to github.com. Create a new github repo (unless you already have one for testing purposes). Navigate to the repo. Click on settings. Select webhooks & Services on the left side. Click on the "Add webhook" buttonAppend "/api/webhooks/incoming/Github" to the  URL of your azure web application (e.g. http://mywebhookspreview.azurewebsites.net/api/webhooks/incoming/Github)You will receive notifications from Github service at this URL.  

12. For the secret, enter the SHA1 key you got from the earlier step

13. Under "Which events would you like to trigger this webhook", select "Let me select individual events". Select the "Issues" checkbox. You will receive notification only when issues are created, assigned, labelled or closed.For this exercise, we will focus on the experience for issue creation.  Click "Update webhook" button to save your Github webhook configuration

14. Optional: to debug the incoming webhook from Github service, place a breakpoint on the ExecuteAsync function of the GithubWebhookHandler.cs file of your Visual Studio project. To do this, open server explorer in Visual Studio, find and right click your web application under App service, and attach the debugger. 

15. Create a new issue for your repo. This will trigger the incoming webhook in your ASP.net application. You should now receive connector card messages for new github issues in your Office 365 group inbox.

**Congratulations,** on completing this exercise! Try building your own connector and submit it to Office 365. Customize the card message experience to show your own brand and style, and the sender avatar will show your logo (instead of "incoming webhook"). Check out dev.outlook.com/connectors for more developer information, code samples and instructions for submitting and listing your connector in the Office 365 Connectors catalog.
