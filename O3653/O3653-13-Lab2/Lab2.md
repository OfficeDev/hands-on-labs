# Find Meeting Times with the Outlook REST API
Learn how to use Outlook REST API  to find the best meeting times between attendees.

## Exercise 1: Create a new project using Azure Active Directory v1 authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v1 Start Project** template, launch your app, log in to your app and generate access tokens.

1. Launch Visual Studio 2015 and select **New>Project**.
  1. Search the installed templates for **Graph** and select the
    **Graph AAD Auth v1 Starter Project** template.
  1. Name the new project **FindMeetingTimesLab** and click **OK**.
  1. Find the **Auth** folder uunder **FindMeetingTimesLab** and open the **AuthHelper.cs** file. 
  1. Search for and replace **https://graph.microsoft.com** with **https://outlook.office.com**. 
1. Press **F5** to compile and launch your new application in the default browser.
  1. The missing NuGet packages should be restored and the app should launch. 
  1. Once the Home page appears, click **Sign in** and login to your Office 365 account.
  1. Review the permissions the application is requesting, and click **Accept**.
  1. Now that you are signed into your application, exercise 1 is complete!
   
## Exercise 2: Access Calendar through REST API

In this exercise, you will build on exercise 1 to connect to the REST API 
endpoint and work with Office 365 and Outlook Calendar. You will be retrieving available meeting times for the signed in user.

## Working with Calendar through REST API  
  
### Create the FindMeetingTime controller

1. Create a new controller to process the requests and send them to the REST API endpoint.
  1. Find the **Controllers** folder under **FindMeetingTimesLab**, right-click it and select **Add>Controller**.
  1. Select **MVC 5 Controller - Empty** and click **Add**.
  1. Change the name of the controller to **FindMeetingTimesController** and click **Add**.

1. **Replace** the following reference to the top of the `CalendarController` class
  ```csharp
  using System;
  using System.Collections.Generic;
  using System.Linq;
  using System.Web;
  using System.Web.Mvc;
  ```

  with the following references
  ```csharp
  using System;
  using System.Web.Mvc;
  using FindMeetingTimesLab.TokenStorage;
  using System.Configuration;
  using System.Threading.Tasks;
  using FindMeetingTimesLab.Auth;
  ```
  
1. Add the following code to the `FindMeetingTimesController` class to generate an access token:

  ```csharp
		[Authorize]
        public async Task<ActionResult> Index()
        {
            string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

            string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");

            AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
            string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));

            if (!string.IsNullOrEmpty((string)TempData["error"]))
            {
                ViewBag.ErrorMessage = (string)TempData["error"];
            }

            return View();
        }
  ```
  
1. Add the following code to the FindMeetingTimesController class to handle initial page load. Add it  after the previous code you added and before the 'return View();' statement.

  ```csharp

            //For first time load, just load the form and the table with no results 
            if (this.Request.HttpMethod == "GET")
            {
                return View();
            }			
  ```
1. Create a new model class.
  1. Find the **Models** folder under **FindMeetingTimesLab**, right-click it and select **Add>New Item**.
  1. Select **Visual C# -> Code** and click **Class** in the middle pane.
  1. Change the name to **MeetingTimeCandidate.cs** and click **Add**.

1. Open the **MeetingTimeCandidate.cs** file and replace the content with the below.

	```csharp
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Web;

	namespace FindMeetingTimesLab.Models
	{
		public class MeetingTimeCandidate
		{
			public DateTime StartDate { get; set; }
			public DateTime EndDate { get; set; }
			public DateTime StartTime { get; set; }
			public DateTime EndTime { get; set; }
			public int Confidence { get; set; }
			public int Score { get; set; }
			public string LocationDisplayName { get; set; }
			public string LocationAddress { get; set; }
			public string LocationCoordinates { get; set; }
		}
	}
	```
		
1. Create a new navigation link in the navigation bar for displaying Meeting Times
  1. Find the **Views** folder under **FindMeetingTimesLab**.
  1. Open **Shared -> _Layout.cshtml**.
  1. Change the name - Replace **Graph and AAD Auth Starter** with **Find Meeting Times Starter**. You would need to do this at 3 places in the file. 
  1. In the `<body>` section, find the place where the labels for the tabbed experience are listed. They will look like `<li>@Html.ActionLink("Home", "Index", "Home")</li>`. 
  1. Add the following to that section
  ```asp
	    <li>@Html.ActionLink("FindMeetingTimes", "Index", "FindMeetingTimes")</li>
  ```

1. Add a View for FindMeetingTimes
  1. Find the **Views** folder under **FindMeetingTimesLab**.
  1. Right click on the **FindMeetingTimes** folder and select **Add -> View**.
  1. Change the name to **Index**.
  1. Select **List** in the **Template** dropdown. 
  1. In the **Model class** dropdown, select **MeetingTimeCandidate**. 
  1. Click **Add**. 
  1. Open the file you just added and replace the contents with the following: 

  ```asp
	@model IEnumerable<FindMeetingTimesLab.Models.MeetingTimeCandidate>

	@{
		ViewBag.Title = "FindMeetingTimes";
	}

	<h2>FindMeetingTimes</h2>

	<div class="table-responsive">
		<table id="calendarTable" class="table table-striped table-bordered">
			<thead>
				<tr>
					<th>StartDate</th>
					<th>StartTime</th>
					<th>EndDate</th>
					<th>EndTime</th>
					<th>Confidence</th>
					<th>Score</th>
					<th>Location Name</th>
					<th>Location Address</th>
					<th>Location Coordinates</th>
				</tr>
			</thead>
			<tbody>
           
			</tbody>
		</table>
	</div>
  ```

1. Press **F5** to compile and launch your application in the default browser.
  1. Once the Home page appears, click **Sign in** and login to your Office 365 account.
  1. Now that you are signed into your application, click on the **FindMeetingTimes** tab. 
  1. Verify that you see the empty table with the correct headings. 
   

### Work with FindMeetingTimes API 

1. Add a button to call **FindMeetingTimes** API.
  1. Open the **Index.cshtml** file you just created and add the following code. Add it above the table you created previously. 

  ```asp
	<div class="panel panel-default">
		<div class="panel-body">
			<form class="form-inline" action="/FindMeetingTimes/Index" method="post">
				<input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
				<button type="submit" class="btn btn-default">Find Meeting Times</button>
			</form>
		</div>
	</div>
  ```
   1. Add the following in the `<tbody>` section of the `<table>` element
	  ```asp
		@if (Model != null)
		{
			foreach (var meetingTimeCandidate in Model)
			{
				<tr>
					<td>
						@meetingTimeCandidate.StartDate
					</td>
					<td>
						@meetingTimeCandidate.StartTime
					</td>                        
					<td>
						@meetingTimeCandidate.EndDate
					</td>
					<td>
						@meetingTimeCandidate.EndTime
					</td>                 
					<td>
						@meetingTimeCandidate.Confidence
					</td>
					<td>
						@meetingTimeCandidate.Score
					</td>
					<td>
						@meetingTimeCandidate.LocationDisplayName
					</td>
					<td>
						@meetingTimeCandidate.LocationAddress
					</td>            
					<td>
						@meetingTimeCandidate.LocationCoordinates
					</td>            
				</tr>
			}
		}
	 ```
   
	1. Add a Graph Helper file.
	  1. Right click on **FindMeetingTimesLab** project and select **Add -> New Item**.
	  1. Select **Visual C# -> Code** and click **Class** in the middle pane.
	  1. Change the name to **GraphHelper.cs** and click **Add**.
	  1. Open the **GraphHelper.cs** file you just created and add replace the contents with the following
	  ```csharp
		using System;
		using System.Collections.Generic;
		using System.Linq;
		using System.Web;
		using System.Threading.Tasks;
		using System.Net.Http;
		using Newtonsoft.Json.Linq;
		using FindMeetingTimesLab.Models;
		using Newtonsoft.Json;
		using System.Text;

		namespace FindMeetingTimesLab
		{
			public class GraphHelper
			{
				// Used to set the base API endpoint, e.g. "https://outlook.office.com/api/beta"
				public string apiEndpoint { get; set; }
				public string anchorMailbox { get; set; }

				public GraphHelper()
				{
					apiEndpoint = "https://outlook.office.com/api/beta";
					anchorMailbox = string.Empty;
				}
			}
		}			  
	 ```

1. Add a method in the **GraphHelper** class to Make the API call. 
   ```csharp
    public async Task<HttpResponseMessage> MakeGraphApiCall(string method, string token, string apiUrl, string userEmail, string payload, Dictionary<string, string> preferHeaders)
    {
        using (var httpClient = new HttpClient())
        {
            var request = new HttpRequestMessage(new HttpMethod(method), apiUrl);

            // Headers
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("FindMeetingTimesLab", "beta"));
            request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
            request.Headers.Add("return-client-request-id", "true");
            request.Headers.Add("X-AnchorMailbox", userEmail);

            if (preferHeaders != null)
            {
                foreach (KeyValuePair<string, string> header in preferHeaders)
                {
                    request.Headers.Add("Prefer", string.Format("{0}=\"{1}\"", header.Key, header.Value));
                }
            }

            // Content
            if ((method.ToUpper() == "POST" || method.ToUpper() == "PATCH") &&
                !string.IsNullOrEmpty(payload))
            {
                request.Content = new StringContent(payload);
                request.Content.Headers.ContentType.MediaType = "application/json";
            }

            var apiResult = await httpClient.SendAsync(request);
            return apiResult;
        }
    }			   
	```

1. Add a method in the **GraphHelper** class to Get Meeting times. 
   ```csharp
	public async Task<object> GetMeetingTimes(string token, string userEmail, string payload)
    {
        string findMeetingTimesEndpoint = this.apiEndpoint + "/Me/FindMeetingTimes";

        //var jsonPayload = await Task.Run(() => JsonConvert.SerializeObject(payload));

        var result = await MakeGraphApiCall("POST", token, findMeetingTimesEndpoint, userEmail, payload, null);

        var response = await result.Content.ReadAsStringAsync();

        JObject responseJson = JObject.Parse(response);
        JArray eventJson = (JArray)responseJson["value"];

        List<MeetingTimeCandidate> meetingTimes = new List<MeetingTimeCandidate>();
        foreach (var e in eventJson)
        {
            MeetingTimeCandidate nextItem = new MeetingTimeCandidate();

            //add all the values
            nextItem.StartDate = DateTime.Parse((string)e["MeetingTimeSlot"]["Start"]["Date"]);
            nextItem.StartTime = DateTime.Parse((string)e["MeetingTimeSlot"]["Start"]["Time"]);
            nextItem.EndDate = DateTime.Parse((string)e["MeetingTimeSlot"]["End"]["Date"]);
            nextItem.EndTime = DateTime.Parse((string)e["MeetingTimeSlot"]["End"]["Time"]);
            nextItem.Confidence = int.Parse((string)e["Confidence"]);
            nextItem.Score = int.Parse((string)e["Score"]);
            if (e["MeetingTimeSlot"]["Location"] != null)
            {
                nextItem.LocationDisplayName = (string)e["MeetingTimeSlot"]["Location"]["Time"];
                nextItem.LocationAddress = BuildAddressString(e["MeetingTimeSlot"]["Location"]["Address"]);
                nextItem.LocationCoordinates = BuildCoordinatesString(e["MeetingTimeSlot"]["Location"]["Coordinates"]);
            }

            meetingTimes.Add(nextItem);
        }

        return meetingTimes;
    }			   
   ```

1. Add a method in the **GraphHelper** class to build the Address string. 
   ```csharp
	public String BuildAddressString(JToken address)
	{
		if (address == null)
		{
			return "null";
		}
		else {
			return String.Format("{0}, {1}, {2}, {3}, {4}",
				(string)address["Street"] == "" ? "<No Street>" : (string)address["Street"],
				(string)address["City"] == "" ? "<No City>" : (string)address["City"],
				(string)address["State"] == "" ? "<No State>" : (string)address["State"],
				(string)address["CountryOrRegion"] == "" ? "<No Country Or Region>" : (string)address["CountryOrRegion"],
				(string)address["PostalCode"] == "" ? "<No Postal Code>" : (string)address["PostalCode"]
				);
		}
	}
   ```

1. Add a method in the **GraphHelper** class to build the coordinates string. 
   ```csharp
	public String BuildCoordinatesString(JToken coordinates)
	{
		if (coordinates == null)
		{
			return "null";
		}
		else
		{
			return String.Format("{0}, {1}", (string)coordinates["Latitude"], (string)coordinates["Longitude"]);
		}
	}		
   ```

1. Add the following code to the **FindMeetingTimesController** class to use **GraphHelper** and call the API. Replace `return View()` at the end of the file with the following

  ```csharp
    try
    {
        var client = new GraphHelper();
        client.anchorMailbox = (string)Session["user_name"];
        ViewBag.UserName = client.anchorMailbox;
		string payload = "";
                
        var results = await client.GetMeetingTimes(accessToken, client.anchorMailbox, payload);

        return View(results);
    }
    catch (Exception ex)
    {
        return RedirectToAction("Index", "Error", new { message = ex.Message });
    }
  ```
1. Run your application by pressing F5. 
1. You should be able to sign-in, click on the FindMeetingTimes tab and see the page with the button and empty table. 
1. Click on the FindMeetingTimes button. The table should load with values. Exercise 2 is complete! 
  
### Get Meeting times for specific attendees 

In this section, you'll add to the code you already created in the previous section.
You will add input parameters to the form and use that as input to the API. 

1. Locate the **GraphHelper.cs** file.
1. Add a method to build the payload string. 
   ```csharp
		public string GeneratePayload(string attendees)
        {
            const string payloadStart = "{";
            const string AttendeesListStart = "\"Attendees\": [";
            const string AttendeeStart = "{\"Type\":\"Required\",\"EmailAddress\":{\"Address\":\"";
            const string AttendeeEnd = "\"}}";
            const string AttendeesListEnd = "]";
            const string payloadEnd = "}";

            StringBuilder payloadBuilder = new StringBuilder(payloadStart);

            //Add all attendees
            string[] attendeeEmails = attendees.Split(',');

            payloadBuilder.Append(AttendeesListStart);
            foreach (var e in attendeeEmails)
            {
                payloadBuilder.Append(AttendeeStart);
                payloadBuilder.Append(e);
                payloadBuilder.Append(AttendeeEnd);
            }
            payloadBuilder.Append(AttendeesListEnd);

            payloadBuilder.Append(payloadEnd);

            return payloadBuilder.ToString();
        }
   ```

  1. Find the **Views/FindMeetingTimes** folder in the project.
  1. Open the **Index.cshtml** file found in the folder.
  1. Locate the part of the file that includes the form at the top of the page. It should look similar to the following code:
   ```asp
		<div class="panel panel-default">
		<div class="panel-body">
			<form class="form-inline" action="/FindMeetingTimes/Index" method="post">
				<input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
				<button type="submit" class="btn btn-default">Find Meeting Times</button>
			</form>
		</div>
		</div>
   ```
  1. Replace this with the following
   ```asp
        <div class="panel panel-default">
            <div class="panel-body">
                <form class="form-inline" action="/FindMeetingTimes/Index" method="post">
                    <div class="form-group">
                        <input type="text" class="form-control" name="Attendees" id="attendees" style="width: 300px;" placeholder="alex@company.com,bob@company.com" />
                    </div>                    
                    <input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
                    <button type="submit" class="btn btn-default">Find Meeting Times</button>
                </form>
            </div>
        </div>
   ```
  1. Find the **FindMeetingTimesController.cs** file and open it. 
  1. Locate the **Index** function definition. It should look like the following
   ```csharp
	 public async Task<ActionResult> Index()
   ```
  1. Replace that line with the following
   ```csharp
        public async Task<ActionResult> Index(string attendees)
   ```
### Run the app

1. Press **F5** to begin debugging.
1. When prompted, login with your Office 365 administrator account.
1. Click the **FindMeetingTimes** link in the navigation bar at the top of the page.
1. Try the app!

**Congratulations dedicated quick start developer!** In this exercise, you created an MVC application that uses REST APIs to view and manage meeting times.
This quick start ends here. You can continue to add more fields to the input. Don't stop here - there's plenty more to explore with the Microsoft Graph endpoint as well.

Next Steps and Additional Resources:

See this training and more on http://dev.office.com/ and http://dev.outlook.com
Learn about and connect to the Microsoft Graph at https://graph.microsoft.io
