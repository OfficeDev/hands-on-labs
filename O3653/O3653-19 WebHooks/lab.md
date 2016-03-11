# Get notified when data changes through Microsoft Graph Webhooks

## TODO  
- changes from beta -> v1.0 path  
- changes to JSON Attributes on model for subscriptionExpirationDateTime -> expirationDateTime and subscriptionId -> id
- change subscriptionExpirationDateTime to DateTimeOffset
- change back to "me/messages" so we get new sent messages too? If not, need to change related verbiage  - think it's less confusing and better guidance to go with inbox and suggest you send a message to yourself. No 404s that way.

## What You'll Learn
In this lab, you'll create an ASP.NET MVC application that subscribes for Microsoft Graph webhooks and receives change notifications. You'll use the Microsoft Graph API to create a subscription, and you'll create a public endpoint that receives change notifications. 

## Overview 
A webhooks subscription allows a client app to receive notifications about mail, events, and contacts from the Microsoft Graph. Microsoft Graph implements a poke-pull model: it sends notifications when changes are made to messages, events, or contacts, and then you query the Microsoft Graph for the details you need. 

## Prerequisites
- Visual Studio 2015 with Update 1
- The Graph AAD Auth v1 Started Project template installed
- An administrator account for an Office 365 tenant. This is required because you'll be using the client credentials of an Azure application that's configured to request admin-level permissions.

## Step 1: Create an ASP.NET MVC application
1. Open Visual Studio and select **File/New/Project**. 

1. In the **New Project** dialog, select **Templates/Visual C#/Graph AAD Auth v1 Starter Project**. If you don't see the template, try searching for *Graph*. The starter project template scaffolds some auth infrastructure for you.

1. Name the new project **GraphWebhooks**, and then click **OK**.  
    
   > NOTE: Make sure you use the exact same name that is specified in these instructions for your Visual Studio project. Otherwise, your namespace name will differ from the one in these instructions and your code will not compile.
 
    ![](images/VSProject.png)

1. Build the solution (**Build/Build Solution**) to restore the NuGet packages required by the project. This should remove all of the solution's initial red squigglies.

1. Open **Tools/Nuget Package Manager/Package Manager Console**, and run the following command. This installs [AspNet.SignalR](http://go.microsoft.com/fwlink/?LinkID=615530), which is used to notify the client to refresh its view.

   ```
Install-Package Microsoft.AspNet.SignalR
   ```

### Configure authorization
This application uses SignalR, which doesn't support ASP.NET session state. So you'll need to reconfigure the **AuthenticationContext** to use the default token cache instead of the **SessionTokenCache** that's provided in the starter template. However, production applications should implement a custom token cache that derives from the ADAL **TokenCache** class. 

1. Open **Startup.cs** in the root directory of the project.

1. Replace the **Configuration** method with the following code.
 

   ```c#
    private async Task OnAuthorizationCodeReceived(AuthorizationCodeReceivedNotification notification)
    {
        // Get the user's object id (used to name the token cache)
        string userObjId = notification.AuthenticationTicket.Identity
            .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

        // Exchange the auth code for a token
        ADAL.ClientCredential clientCred = new ADAL.ClientCredential(appId, appSecret);

        // Create the auth context
        ADAL.AuthenticationContext authContext = new ADAL.AuthenticationContext(
            string.Format(CultureInfo.InvariantCulture, aadInstance, "common", ""),
            false);

        ADAL.AuthenticationResult authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
            notification.Code, notification.Request.Uri, clientCred, "https://graph.microsoft.com");
    }
   ```

1. Open **AccountController.cs** in the Controllers folder
 
1. Replace the **SignOut** method with the following code:


``` C#
        public void SignOut()
        {
            if (Request.IsAuthenticated)
            {
                // Get the user's token cache and clear it
                string userObjId = System.Security.Claims.ClaimsPrincipal.Current
                  .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            }
            // Send an OpenID Connect sign-out request. 
            HttpContext.GetOwinContext().Authentication.SignOut(
              CookieAuthenticationDefaults.AuthenticationType);
            Response.Redirect("/");
        }
```

## Step 2: Set up the ngrok proxy and notification URL data
You must expose a public HTTPS endpoint to create a subscription and receive notifications from Microsoft Graph. While testing, you can use ngrok to temporarily allow messages from Microsoft Graph to tunnel to a port on your local computer. This makes it easier to test and debug webhooks. To learn more about using ngrok, see the [ngrok website](https://ngrok.com/).  

1. In Solution Explorer, select the **GraphWebhooks** project.

1. Copy the **URL** port number from the **Properties** window.  If the **Properties** window isn't showing, choose **View/Properties Window**. 

	![](images/PortNumber.png)

1. [Download ngrok](https://ngrok.com/download) for Windows.  

1. Unzip the package and run ngrok.exe.

1. Replace the two *<port-number>* placeholder values in the following command with the port number you copied, and then run the command in the ngrok console.

   ```
ngrok http <port-number> -host-header=localhost:<port-number>
   ```

	![](images/ngrok1.PNG)

1. Copy the HTTPS URL that's shown in the console. 

	![](images/ngrok2.PNG)

1. In Visual Studio, open the Web.config file in the root directory of the project. Insert the following key in the **appSettings** section, replacing the *<ENTER_YOUR_PROXY_URL>* placeholder value with the URL you just copied.

   ```xml
    <add key="ida:NotificationUrl" value="<ENTER_YOUR_PROXY_URL>/notification/listen" />
   ```

   > NOTE: Keep the console open while testing. If you close it, the tunnel also closes and you'll need to generate a new URL and update the sample.

## Step 3: Configure routing
1. In the **App_Start** folder, open RouteConfig.cs and replace the Default route with the following:

   ```c#
routes.MapRoute(
    name: "Default",
    url: "{controller}/{action}",
    defaults: new { controller = "Subscription", action = "Index" }
);
   ```

## Step 4: Create the Subscription model
In this step you'll create a model that represents a Subscription object. 

1. Right-click the **Models** folder and choose **Add/Class**. 

1. Name the model **Subscription.cs** and click **Add**.

1. Add the following **using** statement. The samples uses the [Json.NET](http://www.newtonsoft.com/json) framework to deserialize JSON responses.

   ```c#
using Newtonsoft.Json;
   ```

1. Replace the **Subscription** class with the following code. This code also includes a view model to display subscription properties in the UI.

   ```c#
    // A webhooks subscription.
    public class Subscription
    {
        // The type of change in the subscribed resource that raises a notification.
        [JsonProperty(PropertyName = "changeType")]
        public string ChangeType { get; set; }

        // The string that MS Graph should send with each notification. Maximum length is 255 characters. 
        // To verify that the notification is from MS Graph, compare the value received with the notification to the value you sent with the subscription request.
        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        // The URL of the endpoint that receives the subscription response and notifications. Requires https.
        [JsonProperty(PropertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        // The resource to monitor for changes.
        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        // The date and time when the webhooks subscription expires.
        // The time is in UTC, and can be up to three days from the time of subscription creation.
        [JsonProperty(PropertyName = "subscriptionExpirationDateTime")]
        public DateTimeOffset? ExpirationDateTime { get; set; }

        // The unique identifier for the webhooks subscription.
        [JsonProperty(PropertyName = "subscriptionId")]
        public string Id { get; set; }
    }

    // The data that displays in the Subscription view.
    public class SubscriptionViewModel
    {
        public Subscription Subscription { get; set; }
    }
   ```

## Step 5: Create the Subscription controller
In this step you'll create a controller that will send a **POST /subscriptions** request to Microsoft Graph on behalf of the signed in user. 

### Create the controller class

1. Right-click the **Controllers** folder and choose **Add/Controller**. 

1. Select **MVC 5 Controller - Empty** and click **Add**.

1. Name the controller **SubscriptionController** and click **Add**.

1. Add the following **using** statements:

   ```c#
using GraphWebhooks.Auth;
using GraphWebhooks.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
   ```

### Create a webhooks subscription

1. Add the **CreateSubscription** method. This gets an access token by calling a helper method, and then adds the token to the HTTP client that sends the **POST /subscriptions** request.

   ```c#
    // Create webhooks subscriptions.
    [Authorize]
    public async Task<ActionResult> CreateSubscription()
    {

        // Get an access token and add it to the client.
        AuthenticationResult authResult = null;
        try
        {
            string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");
            AuthenticationContext authContext = new AuthenticationContext(authority, false);
            ClientCredential credential = new ClientCredential(ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"]);
            authResult = await authContext.AcquireTokenSilentAsync("https://graph.microsoft.com", credential, new UserIdentifier(userObjId, UserIdentifierType.UniqueId));
        }
        catch (Exception ex)
        {
            return RedirectToAction("Index", "Error", new { message = ex.Message, debug = ex.StackTrace });
        }

        HttpClient client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        // Build the request.

        // Send the request and parse the response.

    }
   ```

### Build the POST /subscriptions request

1. Replace the *// Build the request* comment with the following code, which builds the **POST /subscriptions** request. This example uses a random GUID for the client state.

   ```c#
    // Build the request.
    string subscriptionsEndpoint = "https://graph.microsoft.com/beta/subscriptions/";
    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, subscriptionsEndpoint);
    var subscription = new Subscription
    {
        Resource = "me/mailFolders('Inbox')/messages",
        ChangeType = "Created",
        NotificationUrl = ConfigurationManager.AppSettings["ida:NotificationUrl"],
        ClientState = Guid.NewGuid().ToString(),
        ExpirationDateTime = DateTime.UtcNow + new TimeSpan(3, 0, 0, 0)
    };

    string contentString = JsonConvert.SerializeObject(subscription, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
    request.Content = new StringContent(contentString, System.Text.Encoding.UTF8, "application/json");
   ```

This sample creates a subscription for the *me/messages* resource for *Created* change type. See the [docs](http://graph.microsoft.io/en-us/docs/api-reference/beta/resources/subscription) for other supported resources and change types. 

### Send the request and parse the response

1. Replace the *// Send the request and parse the response* comment with the following code. This sends the request, parses the response, and loads the view.

   ```c#
        HttpResponseMessage response = await client.SendAsync(request);
        if (response.IsSuccessStatusCode)
        {

            // Parse the JSON response.
            string stringResult = await response.Content.ReadAsStringAsync();
            SubscriptionViewModel viewModel = new SubscriptionViewModel
            {
                Subscription = JsonConvert.DeserializeObject<Subscription>(stringResult)
            };

            // This app temporarily stores the current subscription ID, refreshToken and client state. 
            // These are required so the NotificationController, which is not authenticated can retrieve an access token keyed from subscription id
            // Production applications typically use some method of persistent storage.
            HttpRuntime.Cache.Insert("subscriptionId_" + viewModel.Subscription.Id, 
                Tuple.Create(viewModel.Subscription.ClientState, authResult.RefreshToken), null, DateTime.MaxValue, new TimeSpan(24, 0, 0), System.Web.Caching.CacheItemPriority.NotRemovable, null);

            // Save the latest subscription ID, so we can delete it later and filter teh view on it.
            Session["SubscriptionId"] = viewModel.Subscription.Id;
            return View("Subscription", viewModel);
        }
        else
        {
            return RedirectToAction("Index", "Error", new { message = response.StatusCode, debug = await response.Content.ReadAsStringAsync() });
        }
   ```

### Build the DELETE /subscriptions/id request

1. Add the **DeleteSubscription** method. This deletes the current subscription and signs the user out.

   ```
    // Delete the current webhooks subscription and sign the user out.
    [Authorize]
    public async Task<ActionResult> DeleteSubscription()
    {
        string subscriptionId = (string) Session["SubscriptionId"];

        if (!string.IsNullOrEmpty(subscriptionId))
        {
            string serviceRootUrl = "https://graph.microsoft.com/beta/subscriptions/";

            // Get an access token and add it to the client.
            string accessToken;
            try
            {
                string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
                string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");
                AuthenticationContext authContext = new AuthenticationContext(authority, false);
                ClientCredential credential = new ClientCredential(ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"]);
                AuthenticationResult authResult = await authContext.AcquireTokenSilentAsync("https://graph.microsoft.com", credential, new UserIdentifier(userObjId, UserIdentifierType.UniqueId));

                accessToken = authResult?.AccessToken;

            }
            catch (Exception ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Message, debug = "" });
            }

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Send the 'DELETE /subscriptions/id' request.
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, serviceRootUrl + subscriptionId);
            HttpResponseMessage response = await client.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                return RedirectToAction("Index", "Error", new { message = response.StatusCode, debug = response.Content.ReadAsStringAsync() });
            }
        }
        return RedirectToAction("SignOut", "Account");
    }
   ```

## Step 6: Create the Index and Subscription views
In this step you'll create a view for the app start page and a view that displays the properties of the subscription you create. You'll also edit the Error view to show a description.

### Create the Index view

1. Right-click the **Views/Subscription** folder and choose **Add/View**. 

1. Name the view **Index**. 

1. Select the **Empty (without model)** template, and then click **Add**.

1. In the **Index.cshtml** file that's created, replace the HTML with the following code:

   ```html
<h2>Microsoft Graph Webhooks</h2>

<div>
    <p>You can subscribe to webhooks for specific resources (such as Outlook messages or events) to get notifications about changes to the resource.</p>
    <p>This sample creates a subscription for the <i>me/messages</i> resource and the <i>Created</i> change type. The request body looks like this:</p>
    <code>
        {<br />
        &nbsp;&nbsp;"resource": "me/messages",<br />
        &nbsp;&nbsp;"changeType": "Created",<br />
        &nbsp;&nbsp;"notificationUrl": "https://your-notification-endpoint",<br />
        &nbsp;&nbsp;"clientState": "your-client-state"<br />
        }
    </code>
    <p>See the <a href="http://graph.microsoft.io/en-us/docs/api-reference/beta/resources/subscription">docs</a> for other supported resources and change types.</p>
    <br />
    @using (Html.BeginForm("CreateSubscription", "Subscription"))
    {
        <button type="submit">Create subscription</button>
    }
</div>
   ```
1. At this point, you can run the app (press **F5**) and sign in as an Office 365 administrator. If you click the **Create subscription** button, the call will fail but you can verify that you can sign in and send an HTTP request.

### Create the Subscription view

1. Right-click the **Views/Subscription** folder and choose **Add/View**. 

1. Name the view **Subscription**.

1. Select the **Empty** template, select **SubscriptionViewModel (GraphWebhooks.Models)**, and then click **Add**.

1. In the **Subscription.cshtml** file, update the HTML as follows:

   ```html
    <h2>Subscription</h2>
    <div>
        <table>
            <tr>
                <td>
                    @Html.LabelFor(m => m.Subscription.Resource, htmlAttributes: new { @class = "control-label col-md-2" })
                </td>
                <td>
                    @Model.Subscription.Resource
                </td>
            </tr>
            <tr>
                <td>
                    @Html.LabelFor(m => m.Subscription.ChangeType, htmlAttributes: new { @class = "control-label col-md-2" })
                </td>
                <td>
                    @Model.Subscription.ChangeType
                </td>
            </tr>
            <tr>
                <td>
                    @Html.LabelFor(m => m.Subscription.ClientState, htmlAttributes: new { @class = "control-label col-md-2" })
                </td>
                <td>
                    @Model.Subscription.ClientState
                </td>
            </tr>
            <tr>
                <td>
                    @Html.LabelFor(m => m.Subscription.Id, htmlAttributes: new { @class = "control-label col-md-2" })
                </td>
                <td>
                    @Model.Subscription.Id
                </td>
            </tr>
            <tr>
                <td>
                    @Html.LabelFor(m => m.Subscription.ExpirationDateTime, htmlAttributes: new { @class = "control-label col-md-2" })
                </td>
                <td>
                    @Model.Subscription.ExpirationDateTime
                </td>
            </tr>
        </table>
    </div>
    <br />
    <div>
        @using (Html.BeginForm("LoadView", "Notification"))
        {
            <button type="submit">Watch for notifications</button>
        }
    </div>
   ```

## Step 7: Create the Notification and Message models
In this step you'll create models that represent Notification and Message objects. 

### Create the Notification model

1. Right-click the **Models** folder and choose **Add/Class**. 

1. Name the model **Notification.cs** and click **Add**.

1. Add the following **using** statement. The sample uses the [Json.NET](http://www.newtonsoft.com/json) framework to deserialize JSON responses.

  ```c#
using Newtonsoft.Json;
  ```

1. Replace the **Notification** class with the following code. This also defines a class for the **ResourceData** object. 

  ```c# 
    // A change notification.
    public class Notification
    {
        // The type of change.
        [JsonProperty(PropertyName = "changeType")]
        public string ChangeType { get; set; }

        // The client state used to verify that the notification is from Microsoft Graph. Compare the value received with the notification to the value you sent with the subscription request.
        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        // The endpoint of the resource that changed. For example, a message uses the format ../Users/{user-id}/Messages/{message-id}
        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        // The date and time when the webhooks subscription expires.
        // The time is in UTC, and can be up to three days from the time of subscription creation.
        [JsonProperty(PropertyName = "subscriptionExpirationDateTime")]
        public string SubscriptionExpirationDateTime { get; set; }

        // The unique identifier for the webhooks subscription.
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        // Properties of the changed resource.
        [JsonProperty(PropertyName = "resourceData")]
        public ResourceData ResourceData { get; set; }
    }

    public class ResourceData
    {

        // The ID of the resource.
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        // The OData etag property.
        [JsonProperty(PropertyName = "@odata.etag")]
        public string ODataEtag { get; set; }

        // The OData ID of the resource. This is the same value as the resource property.
        [JsonProperty(PropertyName = "@odata.id")]
        public string ODataId { get; set; }

        // The OData type of the resource: "#Microsoft.Graph.Message", "#Microsoft.Graph.Event", or "#Microsoft.Graph.Contact".
        [JsonProperty(PropertyName = "@odata.type")]
        public string ODataType { get; set; }
    }
  ```

### Create the Message model

1. Right-click the **Models** folder and choose **Add/Class**. 

1. Name the model **Message.cs** and click **Add**.

1. Add the following **using** statement.

  ```c#
using Newtonsoft.Json;
  ```

1. Replace the **Message** class with the following code. This defines the properties of a Message object that will be displayed in the Notification view. 

  ```c# 
    // An Outlook mail message.
    public class Message
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "subject")]
        public string Subject { get; set; }

        [JsonProperty(PropertyName = "bodyPreview")]
        public string BodyPreview { get; set; }

        [JsonProperty(PropertyName = "createdDateTime")]
        public DateTimeOffset CreatedDateTime { get; set; }

        [JsonProperty(PropertyName = "isRead")]
        public Boolean IsRead { get; set; }

        [JsonProperty(PropertyName = "conversationId")]
        public string ConversationId { get; set; }

        [JsonProperty(PropertyName = "changeKey")]
        public string ChangeKey { get; set; }
    }
  ```

## Step 8: Create the Notification controller
In this step you'll create a controller that exposes the notification endpoint. 

### Create the controller class

1. Right-click the **Controllers** folder and choose **Add/Controller**. 

1. Select **MVC 5 Controller - Empty** and click **Add**.

1. Name the controller **NotificationController** and click **Add**.

1. Add the following **using** statements:

  ```c#

using GraphWebhooks.Auth;
using GraphWebhooks.Models;
using GraphWebhooks.SignalR;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
  ```

### Create the notification endpoint

1. Replace the **Notification** class with the following code. This is the callback method you'll register for notifications.

   ```c#
    // The notificationUrl endpoint that's registered with the webhooks subscription.
    [HttpPost]
    public async Task<ActionResult> Listen()
    {

        // Validate the new subscription by sending the token back to MS Graph.
        // This response is required for each subscription.
        if (Request.QueryString["validationToken"] != null)
        {
            var token = Request.QueryString["validationToken"];
            return Content(token, "plain/text");
        }

        // Parse the received notifications.
        else
        {
            try
            {
                var notifications = new Dictionary<string, Notification>();
                using (var inputStream = new System.IO.StreamReader(Request.InputStream))
                {
                    JObject jsonObject = JObject.Parse(inputStream.ReadToEnd());
                    if (jsonObject != null)
                    {

                        // Notifications are sent in a 'value' array.
                        JArray value = JArray.Parse(jsonObject["value"].ToString());
                        foreach (var notification in value)
                        {
                            Notification current = JsonConvert.DeserializeObject<Notification>(notification.ToString());
                            current.ResourceData = JsonConvert.DeserializeObject<ResourceData>(notification["resourceData"].ToString());

                            var subscriptionParams = (Tuple<string, string>)HttpRuntime.Cache.Get("subscriptionId_" + current.SubscriptionId);
                            if (subscriptionParams != null)
                            {
                                // Verify the message is from Microsoft Graph.
                                if (current.ClientState == subscriptionParams.Item1)
                                {
                                    // Just keep the latest notification for each resource.
                                    // No point pulling data more than once.
                                    notifications[current.Resource] = current;
                                }
                            }
                        }
                        if (notifications.Count > 0)
                        {

                            // Query for the changed messages. 
                            await GetChangedMessagesAsync(notifications.Values);
                        }
                    }
                }
                return new HttpStatusCodeResult(200);
            }
            catch (Exception)
            {

                // TODO: Handle the exception.
                // Return a 200 so the service doesn't resend the notification.
                return new HttpStatusCodeResult(200);
            }
        }
    }
   ```

### Get changed messages

1. Add the **GetChangedMessagesAsync** method to the **NotificationController** class. This queries Microsoft Graph for the changed messages.

   ```c#
    // Get information about the changed messages and send to browser via SignalR.
    // A production application would typically queue a background job for reliability.
    public async Task GetChangedMessagesAsync(IEnumerable<Notification> notifications)
    {
        List<Message> messages = new List<Message>();
        string serviceRootUrl = "https://graph.microsoft.com/v1.0/";

        // Get an access token and add it to the client.
        string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common", "");
        AuthenticationContext authContext = new AuthenticationContext(authority);
        ClientCredential credential = new ClientCredential(ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"]);

        HttpClient client = new HttpClient();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        foreach (var notification in notifications)
        {
            var subscriptionParams = (Tuple<string, string>)HttpRuntime.Cache.Get("subscriptionId_" + notification.SubscriptionId);
            string refreshToken = subscriptionParams.Item2;
            AuthenticationResult authResult = await authContext.AcquireTokenByRefreshTokenAsync(refreshToken, credential, "https://graph.microsoft.com");

            // Send the 'GET' request.
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, serviceRootUrl + notification.Resource);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(continueOnCapturedContext: false);

            // Get the messages from the JSON response.
            if (response.IsSuccessStatusCode)
            {
                string stringResult = await response.Content.ReadAsStringAsync();
                var type = notification.ResourceData.ODataType;
                if (type == "#Microsoft.Graph.Message")
                {
                    messages.Add(JsonConvert.DeserializeObject<Message>(stringResult));
                }
            }
        }
        if (messages.Count > 0)
        {
            NotificationService notificationService = new NotificationService();
            notificationService.SendNotificationToClient(messages);
        }
    }
   ```

## Step 9: Set up SignalR

This app uses SignalR to notify the client to refresh its view.

1. Open **Startup.cs** in the root directory of the project.

1. Add the following line to the **Configuration** method.

   ```c#
app.MapSignalR();
   ```

1. Right-click the **GraphWebhooks** project and create a folder named **SignalR**.

1. Right-click the **SignalR** folder and choose **Add/SignalR Hub Class (v2)**. 

1. Name the class **NotificationHub**, and click **OK**. This sample doesn't add any functionality to the hub.

1. Right-click the **SignalR** folder and choose **Add/New Item**. Choose the **SignalR/SignalR Persistent Connection Class (v2)** template.

1. Name the class **NotificationService.cs**, and click **Add**.

1. In **NotificationService**, add the following **using** statement:

   ```c#
using GraphWebhooks.Models;
using System.Threading.Tasks;
   ```

1. Replace the **NotificationService** class with the following code.

   ```c#
    public class NotificationService : PersistentConnection
    {
      public void SendNotificationToClient(List<Message> messages)
        {
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<NotificationHub>();
            if (hubContext != null)
            {
                hubContext.Clients.All.showNotification(messages);
            }
        }
    }
   ```

## Step 10: Create the Notification view
In this step you'll create a view that displays some properties of the changed message. 

1. Right-click the **Views/Subscription** folder and choose **Add/View**. 

1. Name the view **Notification**.

1. Select the **Empty** template, select **Message (GraphWebhooks.Models)**, and then click **Add**.

1. In the **Notification.cshtml* file that's created, replace the content with the following code:

   ```html
@model GraphWebhooks.Models.Message

@{
    ViewBag.Title = "Notification";
}

@section Scripts {
    @Scripts.Render("~/Scripts/jquery.signalR-2.2.0.min.js");
    @Scripts.Render("~/signalr/hubs");

    <script>
    var connection = $.hubConnection();
    var hub = connection.createHubProxy("NotificationHub");
    hub.on("showNotification", function (messages) {
        $.each(messages, function (index, value) {     // Iterate through the message collection
            var message = value;                       // Get current message

            var table = $("<table></table>");
            var header = $("<th>Message " + (index + 1) + "</th>").appendTo(table);

            for (prop in message) {                    // Iterate through message properties
                var property = message[prop];
                var row = $("<tr></tr>");

                $("<td></td>").text(prop).appendTo(row);
                $("<td></td>").text(property).appendTo(row);
                table.append(row);
            }
            $("#message").append(table);
            $("#message").append("<br />");
        });
    });
    connection.start();
    </script>
}
<h2>Messages</h2>
<p>You'll get a notification when your user receives an email. The messages display below.</p>
<br />
<div id="message"></div>
<div>
    @using (Html.BeginForm("DeleteSubscription", "Subscription"))
    {
        <button type="submit">Delete subscription and sign out</button>
    }
</div>
   ```

Congratulations! In this exercise you created an MVC application that subscribes for Microsoft Graph webhooks and receives change notifications! Now you can run the app.

## Step 10: Run the application

1. Make sure that the ngrok console is still running, then press **F5** to begin debugging.

1. Sign in with your Office 365 administrator account.

1. Click the **Create subscription** button. The **Subscription** page loads with information about the subscripton.

1. Click the **Watch for notifications** button.

1. Send an email to your administrator account, or send an email from the administrator account. The **Notification** page displays information about the message.

1. Click the **Delete subscription and sign out** button. 

## Next Steps and Additional Resources:  
- See this training and more on http://dev.office.com/
- Learn about and connect to the Microsoft Graph at https://graph.microsoft.io
