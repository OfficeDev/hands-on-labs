# Microsoft Graph for Mail and Calendar
In this lab, you will use Microsoft Graph to program against an Office 365 and Outlook mailbox with an ASP.NET MVC application.

[//]: # (Change which template based on if using converged auth)

## Prerequisites
1. You must have an Office 365 tenant and Microsoft Azure subscription to
   complete this lab. If you do not have one, the lab for **O3651-7 Setting up
   your Developer environment in Office 365** shows you how to obtain a trial.
1. You must have Visual Studio 2015 with Update 1 installed.
1. You must have the Graph AAD Auth v1 Starter Project template installed.

[//]: # (Remove if doing v1) 

## Exercise 1: Create a new project using Azure Active Directory v2 authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v2 Start Project** template, register a new application
in the developer portal, and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **New**, **Project**.
  1. Search the installed templates for **Graph** and select the
    **Graph AAD Auth v2 Starter Project** template.
  1. Name the new project **MailCalendarWebApp** and click **OK**.
  1. Open the **Web.config** file and find the **appSettings** element. This is where you will need to add your appId and app secret you will generate in the next step.
1. Launch the [Application Registration Portal](https://apps.dev.microsoft.com)
   to register a new application.
  1. Sign into the portal using your Office 365 username and password.
  1. Click **Add an App** and type **Graph MailCalendar Quick Start** for the application name.
  1. Copy the **Application Id** and paste it into the value for **ida:AppId** in your project's **web.config** file.
  1. Under **Application Secrets** click **Generate New Password** to create a new client secret for your app.
  1. Copy the displayed app password and paste it into the value for **ida:AppSecret** in your project's **web.config** file.
  1. Modify the **ida:AppScopes** value to include the required `https://outlook.office.com/mail.readwrite`, `https://outlook.office.com/mail.send` and `https://outlook.office.com/calendars.readwrite`  scopes.

  ```xml
  <configuration>
    <appSettings>
      <!-- ... -->
      <add key="ida:AppId" value="paste application id here" />
      <add key="ida:AppSecret" value="paste application password here" />
      <!-- ... -->
      <!-- Specify scopes in this value. Multiple values should be comma separated. -->
      <add key="ida:AppScopes" value="https://outlook.office.com/mail.readwrite,https://outlook.office.com/mail.send,https://outlook.office.com/calendars.readwrite" />
    </appSettings>
    <!-- ... -->
  </configuration>
  ```
1. Add a redirect URL to enable testing on your localhost.
  1. Right click on **MailCalendarWebApp** and click on **Properties** to open the project properties.
  1. Click on **Web** in the left navigation.
  1. Copy the **Project Url** value.
  1. Back on the Application Registration Portal page, click **Add Platform** and then **Web**.
  1. Paste the value of **Project Url** into the **Redirect URIs** field.
  1. Scroll to the bottom of the page and click **Save**.

1. Press F5 to compile and launch your new application in the default browser.
  1. Once the Graph and AAD v2 Auth Endpoint Starter page appears, click **Sign in** and login to your Office 365 account.
  1. Review the permissions the application is requesting, and click **Accept**.
  1. Now that you are signed into your application, exercise 1 is complete!
   
[//]: # (Remove if doing v2)

## Exercise 1: Create a new project using Azure Active Directory authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v1 Starter Project** template and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **New**, **Project**.
  1. Search the installed templates for **Graph** and select the
    **Graph AAD Auth v1 Starter Project** template.
  1. Name the new project **MailCalendarWebApp** and click **OK**.
   
1. Press F5 to compile and launch your new application in the default browser.
  1. Once the Graph and AAD Auth Endpoint Starter page appears, click **Sign in** and login to your Office 365 adminsitrator account.
  1. Review the permissions the application is requesting, and click **Accept**.
  1. Now that you are signed into your application, exercise 1 is complete!
   
## Exercise 2: Access Mail, Calendar through Microsoft Graph SDK

In this exercise, you will build on exercise 1 to connect to the Microsoft Graph
SDK and work with Office 365 and Outlook Mail, Calendar

## Working with Mail through Microsoft Graph SDK
  
### Create the Mail controller and use the Graph SDK

1. Add a reference to the Microsoft Graph SDK to your project
  1. In the **Solution Explorer** right click on the **MailCalendarWebApp** project and select **Manage NuGet Packages...**.
  1. Click **Browse** and search for **Microsoft.Graph**.
  1. Select the Microsoft Graph SDK and click **Install**.
  
1. Add a reference to the Bootstrap DateTime picker to your project
  1. In the **Solution Explorer** right click on the **MailCalendarWebApp** project and select **Manage NuGet Packages...**.
  1. Click **Browse** and search for **Bootstrap.v3.Datetimepicker.CSS**.
  1. Select Bootstrap.v3.Datetimepicker.CSS and click **Install**.
  1. Open the **App_Start/BundleConfig.cs** file and update the bootstrap script and CSS bundles. Replace these lines:
  
    ```csharp
    bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
              "~/Scripts/bootstrap.js",
              "~/Scripts/respond.js"));

    bundles.Add(new StyleBundle("~/Content/css").Include(
              "~/Content/bootstrap.css",
              "~/Content/site.css"));
    ```
    
    with:
    
    ```csharp
    bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
              "~/Scripts/bootstrap.js",
              "~/Scripts/respond.js",
              "~/Scripts/moment.js",
              "~/Scripts/bootstrap-datetimepicker.js"));

    bundles.Add(new StyleBundle("~/Content/css").Include(
              "~/Content/bootstrap.css",
              "~/Content/bootstrap-datetimepicker.css",
              "~/Content/site.css"));
    ```

1. Create a new controller to process the requests for files and send them to Graph API.
  1. Find the **Controllers** folder under **MailCalendarWebApp**, right click on it and select **Add** then **Controller**.
  1. Select **MVC 5 Controller - Empty** and click **Add**.
  1. Change the name of the controller to **MailController** and click **Add**.

1. **Add** the following reference to the top of the `MailController` class.

  ```csharp
  using System;
  using System.Collections.Generic;
  using System.Linq;
  using System.Web;
  using System.Web.Mvc;
  using System.Configuration;
  using System.Threading.Tasks;
  using Microsoft.Graph;
  using MailCalendarWebApp.Auth;
  using MailCalendarWebApp.TokenStorage;
  using Newtonsoft.Json;
  using System.IO;
  ```
  
1. Add the following code to the `MailController` class to initialize a new
   `GraphServiceClient` and generate an access token for the Graph API:

  ```csharp
  private GraphServiceClient GetGraphServiceClient()
  {
    string userObjId = System.Security.Claims.ClaimsPrincipal.Current
      .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
    SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

    string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common", "");

    AuthHelper authHelper = new AuthHelper(
      authority,
      ConfigurationManager.AppSettings["ida:AppId"],
      ConfigurationManager.AppSettings["ida:AppSecret"],
      tokenCache);

    // Request an accessToken and provide the original redirect URL from sign-in
    GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
    {
      string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
      request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
    }));

    return client;
  }
  ```
  
### Work with Mails
  
1. Add the following code to the `MailController` class to get all mails from your mailbox.

  ```csharp
        // GET: Me/Messages
        [Authorize]
        public async Task<ActionResult> Index(int? pageSize, string nextLink)
        {
            if (!string.IsNullOrEmpty((string)TempData["error"]))
            {
                ViewBag.ErrorMessage = (string)TempData["error"];
            }

            pageSize = pageSize ?? 25;

            var client = GetGraphServiceClient();
            var request = client.Me.MailFolders.Inbox.Messages.Request().Top(pageSize.Value);
            if (!string.IsNullOrEmpty(nextLink))
            {
                request = new MessagesCollectionRequest(nextLink, client, null);
            }

            try
            {
                var results = await request.GetAsync();

                ViewBag.NextLink = null == results.NextPageRequest ? null :
                  results.NextPageRequest.GetHttpRequestMessage().RequestUri;

                return View(results);
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

  ```
  
1. Add the following code to the `MailController` class to display details of a mail.

  ```csharp
        // GET: Message/Detail?messageId=<id>
        [Authorize]
        public async Task<ActionResult> Detail(string messageId)
        {
            var client = GetGraphServiceClient();

            var request = client.Me.Messages[messageId].Request();

            try
            {
                var result = await request.GetAsync();
                if (result.Body.ContentType.HasValue)
                {
                    if (result.Body.ContentType == BodyType.html)
                    {
                        string textStr = StripHTML(result.Body.Content);
                        result.Body.Content = textStr;
                    }
                }
                    
                return View(result);
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }
        }
  ```
  
1. Add the following code to the `MailController` class to send a new mail.

  ```csharp
        // POST Messages/SendMail
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> SendMail(string messageId, string recipients,  string subject, string body)
        {
            if (string.IsNullOrEmpty(subject) || string.IsNullOrEmpty(recipients))
            {
                TempData["error"] = "Please fill in all fields";
            }
            else
            {
                var client = GetGraphServiceClient();
                ItemBody CurrentBody = new ItemBody();
                CurrentBody.Content = (string.IsNullOrEmpty(body) ? "" : body);
                EmailAddress mailAddress = new EmailAddress();
                mailAddress.Address = "sarad@tr22rest.onmicrosoft.com";
                Recipient mailRecipient = new Recipient();
                mailRecipient.EmailAddress = mailAddress;
                List<Recipient> mailRecipients = new List<Recipient>();
                mailRecipients.Add (mailRecipient);

                Message newMessage = new Message()
                {
                    Subject = subject,
                    Body = CurrentBody,
                    ToRecipients = mailRecipients
                };
                var request = client.Me.SendMail(newMessage, true).Request();

                try
                {
                    await request.PostAsync();
                }
                catch (ServiceException ex)
                {
                    TempData["error"] = ex.Error.Message;
                }
            }

            return RedirectToAction("Index", new { messageId = messageId });
        }
  ```
  
1. Add the following code to the `MailController` class to reply to a mail.

  ```csharp
        // POST Messages/<<ID>>/Reply
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> Reply(string messageId, string comment)
        {
            var client = GetGraphServiceClient();
                
            var request = client.Me.Messages[messageId].Reply(comment).Request();

            try
            {
                await request.PostAsync();
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

            return RedirectToAction("Detail", new { messageId = messageId });
        }
  ```
  
1. Add the following code to the `MailController` class to reply all to a mail.

  ```csharp
        // POST Messages/<<ID>>/ReplyAll
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> ReplyAll(string messageId, string comment)
        {
            var client = GetGraphServiceClient();

            var request = client.Me.Messages[messageId].ReplyAll(comment).Request();

            try
            {
                await request.PostAsync();
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

            return RedirectToAction("Detail", new { messageId = messageId });
        }
  ```
  
  1. Add the following code to the `MailController` class to forward a mail.

  ```csharp
        // POST Messages/<<ID>>/Forward
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> Forward(string messageId, string comment, string recipients)
        {
            var client = GetGraphServiceClient();
            EmailAddress mailAddress = new EmailAddress();
            mailAddress.Address = "zrinkam@tr22rest.onmicrosoft.com";
            Recipient mailRecipient = new Recipient();
            mailRecipient.EmailAddress = mailAddress;

            var request = client.Me.Messages[messageId].Forward(comment).Request();

            try
            {
                await request.PostAsync();
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

            return RedirectToAction("Detail", new { messageId = messageId });
        }
  ```
  
### Create the MailList view

In this section you'll wire up the Controller you created in the previous section
to an MVC view that will display all your mails and allow you to send a new mail.

1. Locate the **Views/Shared** folder in the project.
1. Open the **_Layout.cshtml** file found in the **Views/Shared** folder.
  1. Locate the part of the file that includes a few links at the top of the
      page. It should look similar to the following code:

  ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Graph API", "Graph", "Home")</li>
    </ul>
  ```

  1. Update that navigation to replace the "Graph API" link with "Outlook Mail API"
      and connect this to the MailController you just created.

  ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Outlook Mail API", "Index", "Mail")</li>
    </ul>
  ```
1. Create a new **View** for MailList.
  1. Expand the **Views** folder in **MailCalendarWebApp**. Right-click **Mail** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Index.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Mail/Index.cshtml** with the following:
  
  ```asp
@model IEnumerable<Microsoft.Graph.Message>
@{ ViewBag.Title = "Index"; }
<h2>Inbox</h2>
@section scripts {
    <script type="text/javascript">
$(function () {
    $('#start-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
    $('#end-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
});
    </script>
}
<div class="row" style="margin-top:50px;">
    <div class="col-sm-12">
        @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
        {
            <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
        }
        <div class="panel panel-default">
            <div class="panel-body">
                <form class="form-inline" action="/Mail/SendMail" method="post">
                    <div class="form-group">
                        <input type="text" class="form-control" name="recipients" id="recipients" placeholder="To" />
                    </div>
                    <div class="form-group">
                        <input type="text" class="form-control" name="subject" id="subject" placeholder="Subject" />
                    </div>
                    <div class="form-group">
                        <input type="text" class="form-control" name="body" id="body" placeholder="body" />
                    </div>
                    <input type="hidden" name="messageId" value="@Request.Params["messageId"]" />
                    <button type="submit" class="btn btn-default">Send Mail</button>
                </form>
            </div>
        </div>
        <div class="table-responsive">
            <table id="MailTable" class="table table-striped table-bordered">
                <thead>
                    <tr>
                        <th>From</th>
                        <th><p>Subject</p><p>Body Preview</p></th>
                        <th>Received</th>
                        <th>Has Attachments</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach (var MailMessage in Model)
                    {
                        <tr>
                            <td>
                                @MailMessage.From.EmailAddress.Name
                            </td>
                            <td>
                                @{
                                    RouteValueDictionary idVal = new RouteValueDictionary();
                                    idVal.Add("messageId", MailMessage.Id);
                                    if (null != @MailMessage.IsRead)
                                    {
                                        if ((bool)MailMessage.IsRead)
                                        {
                                            <p>
                                               @Html.ActionLink(MailMessage.Subject, "Detail", idVal)
                                            </p>
                                        }

                                        else
                                        {
                                            <p>
                                                <b>
                                                    @Html.ActionLink(MailMessage.Subject, "Detail", idVal)
                                                </b>
                                            </p>
                                        }
                                    }
                                }
                                @MailMessage.BodyPreview
                            </td>
                            <td>
                                @{
                                    if (null != MailMessage.ReceivedDateTime)
                                    {
                                        @MailMessage.ReceivedDateTime
                                    }
                                }
                            </td>
                            <td>
                                @{
                                    if (null != MailMessage.HasAttachments)
                                    {
                                        @((bool)MailMessage.HasAttachments ? "Yes" : "No")
                                    }
                                }
                            </td>
                        </tr>
                    }
                </tbody>
            </table>
        </div>
        <div class="btn btn-group-sm">
            @{
                Dictionary<string, object> attributes = new Dictionary<string, object>();
                attributes.Add("class", "btn btn-default");

                if (null != ViewBag.NextLink)
                {
                    RouteValueDictionary routeValues = new RouteValueDictionary();
                    routeValues.Add("nextLink", ViewBag.NextLink);
                    @Html.ActionLink("Next Page", "Index", "Mail", routeValues, attributes);
                }
            }
        </div>

    </div>
</div>
  ```
  
1. Create a new **View** for mail detail.
  1. Expand the **Views** folder in **MailCalendarWebApp**. Right-click **Mail** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Detail.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Mail/Detail.cshtml** with the following:
  
  ```asp
@model Microsoft.Graph.Message
@{ ViewBag.Title = "Detail"; }
<h2>@Model.Subject</h2>
@section scripts {
    <script type="text/javascript">
$(function () {
    $('#start-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
    $('#end-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
});
    </script>
}
<div class="row" style="margin-top:50px;">
    <div class="col-sm-12">
        @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
        {
            <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
        }

        <div class="table-responsive">
            <table id="messageTable" class="table table-striped table-bordered">
                <tbody>
                    <tr>
                        <td class="auto-style8">From:</td>
                        <td>
                            @Model.From.EmailAddress.Name
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style8">To:</td>
                        <td>
                            @{
                                string toRecipients = "";

                                foreach (var To in Model.ToRecipients)
                                {
                                    toRecipients = toRecipients + To.EmailAddress.Name + "; ";
                                }
                            }
                            @toRecipients
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style8">Cc:</td>
                        <td>
                            @{
                                string ccRecipients = "";

                                foreach (var To in Model.CcRecipients)
                                {
                                    ccRecipients = ccRecipients + To.EmailAddress.Name + "; ";
                                }
                            }
                            @ccRecipients
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style8">Received:</td>
                        <td>
                            @{
                                if (null != Model.ReceivedDateTime)
                                {
                                    @Model.ReceivedDateTime
                                }
                            }
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style8">Has Attachments:</td>
                        <td>
                            @{
                                if (null != Model.HasAttachments)
                                {
                                    @((bool)Model.HasAttachments ? "Yes" : "No")
                                }
                            }
                        </td>
                    </tr>
                    <tr>
                        <td class="auto-style8">Body:</td>
                        <td>
                            <pre>
                                @{
                                    @Model.Body.Content
                                }
                            </pre>
                        </td>
                    </tr>
                    <tr>
                </tbody>
            </table>
        </div>
        <div class="panel panel-default">
            <div class="panel-body">
                <table>
                    <tbody>
                        <tr>
                            <form class="form-inline" action="/Mail/Reply" method="post">
                                <td class="auto-style2">
                                   <input type="text" class="auto-style4" name="comment" id="comment" placeholder="Comment" />
                                </td>

                                <td class="auto-style3">
                                    <input type="hidden" name="messageId" value="@Request.Params["messageId"]" />
                                    <button type="submit" name="Reply" class="btn btn-default">Reply</button>
                                </td>
                            </form>
                        </tr>
                        <br />
                        <tr>
                            <form class="form-inline" action="/Mail/ReplyAll" method="post">
                                <td class="auto-style1">
                                    <input type="text" class="auto-style5" name="comment" id="comment" placeholder="Comment" />
                                </td>

                                <td>
                                    <input type="hidden" name="messageId" value="@Request.Params["messageId"]" />
                                    <button type="submit" name="Reply All" class="btn btn-default">Reply All</button>
                                </td>
                            </form>
                        </tr>
                        <tr>
                            <form class="form-inline" action="/Mail/Forward" method="post">
                            <td class="auto-style1">
                                <input type="text" class="auto-style6" name="recipients" id="recipients" placeholder="To" />
                            </td>
                            <td class="auto-style1">
                                <input type="text" class="auto-style7" name="comment" id="comment" placeholder="Comment" />
                            </td>
                            <td>
                                <input type="hidden" name="messageId" value="@Request.Params["messageId"]" />
                                <button type="submit" name="Forward" class="btn btn-default">Forward</button>
                            </td>
                        </tr>
                            </form>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
  ```


## Working with Calendar through Microsoft Graph SDK

### Create the Calendar controller and use the Graph SDK

// 1. Add a reference to the Microsoft Graph SDK to your project
//   1. In the **Solution Explorer** right click on the **MailCalendarWebApp** project and select **Manage NuGet Packages...**.
//   1. Click **Browse** and search for **Microsoft.Graph**.
//   1. Select the Microsoft Graph SDK and click **Install**.
//   
// 1. Add a reference to the Bootstrap DateTime picker to your project
//   1. In the **Solution Explorer** right click on the **MailCalendarWebApp** project and select **Manage NuGet Packages...**.
//   1. Click **Browse** and search for **Bootstrap.v3.Datetimepicker.CSS**.
//   1. Select Bootstrap.v3.Datetimepicker.CSS and click **Install**.
//   1. Open the **App_Start/BundleConfig.cs** file and update the bootstrap script and CSS bundles. Replace these lines:
//   
//     ```csharp
//     bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
//               "~/Scripts/bootstrap.js",
//               "~/Scripts/respond.js"));
// 
//     bundles.Add(new StyleBundle("~/Content/css").Include(
//               "~/Content/bootstrap.css",
//               "~/Content/site.css"));
//     ```
//     
//     with:
//     
//     ```csharp
//     bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
//               "~/Scripts/bootstrap.js",
//               "~/Scripts/respond.js",
//               "~/Scripts/moment.js",
//               "~/Scripts/bootstrap-datetimepicker.js"));
// 
//     bundles.Add(new StyleBundle("~/Content/css").Include(
//               "~/Content/bootstrap.css",
//               "~/Content/bootstrap-datetimepicker.css",
//               "~/Content/site.css"));
//     ```

1. Create a new controller to process the requests for Outlook Calendar and send them to Graph API.
  1. Find the **Controllers** folder under **MailCalendarWebApp**, right click on it and select **Add** then **Controller**.
  1. Select **MVC 5 Controller - Empty** and click **Add**.
  1. Change the name of the controller to **CalendarController** and click **Add**.

1. **Add** the following reference to the top of the `CalendarController` class.

  ```csharp
  using System;
  using System.Collections.Generic;
  using System.Linq;
  using System.Web;
  using System.Web.Mvc;
  using System.Configuration;
  using System.Threading.Tasks;
  using Microsoft.Graph;
  using MailCalendar.Auth;
  using MailCalendar.TokenStorage;
  using Newtonsoft.Json;
  using System.IO;
  ```
  
1. Add the following code to the `CalendarController` class to initialize a new
   `GraphServiceClient` and generate an access token for the Graph API:

  ```csharp
  private GraphServiceClient GetGraphServiceClient()
  {
    string userObjId = System.Security.Claims.ClaimsPrincipal.Current
      .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
    SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

    string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common", "");

    AuthHelper authHelper = new AuthHelper(
      authority,
      ConfigurationManager.AppSettings["ida:AppId"],
      ConfigurationManager.AppSettings["ida:AppSecret"],
      tokenCache);

    // Request an accessToken and provide the original redirect URL from sign-in
    GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
    {
      string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
      request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
    }));

    return client;
  }
  ```
  
### Work with EventList
  
1. Add the following code to the `CalenderController` class to get all events in your mailbox.

  ```csharp
        // GET: Me/Events
        [Authorize]
        public async Task<ActionResult> Index(int? pageSize, string nextLink)
        {
            if (!string.IsNullOrEmpty((string)TempData["error"]))
            {
                ViewBag.ErrorMessage = (string)TempData["error"];
            }

            pageSize = pageSize ?? 10;

            var client = GetGraphServiceClient();

            DateTime start = DateTime.Today;
            DateTime end = start.AddDays(6);
            List<Option> viewOptions = new List<Option>();
            viewOptions.Add(new QueryOption("startDateTime",
              start.ToUniversalTime().ToString("s", System.Globalization.CultureInfo.InvariantCulture)));
            viewOptions.Add(new QueryOption("endDateTime",
              end.ToUniversalTime().ToString("s", System.Globalization.CultureInfo.InvariantCulture)));

            var request = client.Me.CalendarView.Request(viewOptions).Top(pageSize.Value);
            if (!string.IsNullOrEmpty(nextLink))
            {
                request = new CalendarViewCollectionRequest(nextLink, client, null);
            }

            try
            {
                var results = await request.GetAsync();

                ViewBag.NextLink = null == results.NextPageRequest ? null :
                  results.NextPageRequest.GetHttpRequestMessage().RequestUri;

                return View(results);
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }
  ```
  
1. Add the following code to the `CalenderController` class to display details of an event.

  ```csharp
        // GET: Event/Detail?eventId=<id>
        [Authorize]
        public async Task<ActionResult> Detail(string eventId)
        {
            var client = GetGraphServiceClient();

            var request = client.Me.Events[eventId].Request();

            try
            {
                var result = await request.GetAsync();

                if (result.Body.ContentType.HasValue)
                {
                    if (result.Body.ContentType == BodyType.html)
                    {
                        string textStr = StripHTML(result.Body.Content);
                        result.Body.Content = textStr;
                    }
                }

                return View(result);
            }
            catch (ServiceException ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }
        }
  ```
  
1. Add the following code to the `CalendarController` class to add a new event in the calendar.

  ```csharp
  // POST: me/events?eventId=<id>&subject=<text>&start=<text>&end=<text>&location=<text>
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> AddEvent(string eventId, string subject, string body, string start, string end, string location)
        {
            if (string.IsNullOrEmpty(subject) || string.IsNullOrEmpty(start)
              || string.IsNullOrEmpty(end) || string.IsNullOrEmpty(location))
            {
                TempData["error"] = "Please fill in all fields";
            }
            else
            {
                var client = GetGraphServiceClient();
                
                var request = client.Me.Events.Request();

                ItemBody CurrentBody = new ItemBody();
                CurrentBody.Content = (string.IsNullOrEmpty(body) ? "" : body);
                Event newEvent = new Event()
                {
                    Subject = subject,
                    Body = CurrentBody,
                    Start = new DateTimeTimeZone() { DateTime = start, TimeZone = "UTC" },
                    End = new DateTimeTimeZone() { DateTime = end, TimeZone = "UTC" },
                    Location = new Location() { DisplayName = location }
                };

                try
                {
                    await request.AddAsync(newEvent);
                }
                catch (ServiceException ex)
                {
                    TempData["error"] = ex.Error.Message;
                }
            }

            return RedirectToAction("Index", new { eventId = eventId });
  ```
  
1. Add the following code to the `CalendarController` class to Accept an event.

  ```csharp
  // POST: me/events/<<ID>>/Accept
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> Accept(string eventId)
        {
            var client = GetGraphServiceClient();

            var request = client.Me.Events[eventId].Accept().Request();

            try
            {
                await request.PostAsync();
        }
        catch (ServiceException ex)
            {
                TempData["error"] = ex.Error.Message;
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

            return RedirectToAction("Detail", new { eventId = eventId });
        }

  ```
   
1. Add the following code to the `CalendarController` class to TentativelyAccept an event.

  ```csharp
  // POST: me/events/<<ID>>/TentativelyAccept
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> Tentative(string eventId)
        {
            var client = GetGraphServiceClient();

            var request = client.Me.Events[eventId].TentativelyAccept().Request();

            try
            {
                await request.PostAsync();
            }
            catch (ServiceException ex)
            {
                TempData["error"] = ex.Error.Message;
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

            return RedirectToAction("Detail", new { eventId = eventId });
        }
  ```
1. Add the following code to the `CalendarController` class to Decline an event.

  ```csharp
  // POST: me/events/<<ID>>/Decline
        [Authorize]
        [HttpPost]
        public async Task<ActionResult> Decline(string eventId)
        {
            var client = GetGraphServiceClient();

            var request = client.Me.Events[eventId].Decline().Request();

            try
            {
                await request.PostAsync();
            }
            catch (ServiceException ex)
            {
                TempData["error"] = ex.Error.Message;
                return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
            }

            return RedirectToAction("Index");
        }
  ```
   
  
### Create the CalendarList view

In this section you'll wire up the CalendarController you created in the previous section
to an MVC view that will display the events in your calendar and allow you to add an event to it.

1. Locate the **Views/Shared** folder in the project.
1. Open the **_Layout.cshtml** file found in the **Views/Shared** folder.
  1. Locate the part of the file that includes a few links at the top of the
      page. It should look similar to the following code:

  ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Outlook Mail API", "Index", "Mail")</li>
    </ul>
  ```

  1. Update that navigation to add the "Outlook Calendar API" link with "Calendar"
      and connect this to the controller you just created.

  ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Outlook Mail API", "Index", "Mail")</li>
        <li>@Html.ActionLink("Outlook Calendar API", "Index", "Calendar")</li>
    </ul>
  ```
1. Create a new **View** for CalendarList.
  1. Expand the **Views** folder in **MailCalendarWebApp**. Right-click **Calendar** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Index.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Calendar/Index.cshtml** with the following:
  
  ```asp
@model IEnumerable<Microsoft.Graph.Event>
@{ ViewBag.Title = "Index"; }
<h2>Calendar (Next 7 Days)</h2>
@section scripts {
    <script type="text/javascript">
$(function () {
    $('#start-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
    $('#end-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
});
    </script>
}
<div class="row" style="margin-top:50px;">
    <div class="col-sm-12">
        @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
        {
            <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
        }
        <div class="panel panel-default">
            <div class="panel-body">
                <form class="form-inline" action="/Calendar/AddEvent" method="post">
                    <div class="form-group">
                        <input type="text" class="form-control" name="subject" id="subject" placeholder="Subject" />
                    </div>
                    <div class="form-group">
                        <input type="text" class="form-control" name="body" id="body" placeholder="body" />
                    </div>
                    <div class="form-group">
                        <div class="input-group date" id="start-picker">
                            <input type="text" class="form-control" name="start" id="start" placeholder="Start Time (UTC)" />
                            <span class="input-group-addon">
                                <span class="glyphicon glyphicon-calendar"></span>
                            </span>
                        </div>
                    </div>
                    <div class="form-group">
                        <div class="input-group date" id="end-picker">
                            <input type="text" class="form-control" name="end" id="end" placeholder="End Time (UTC)" />
                            <span class="input-group-addon">
                                <span class="glyphicon glyphicon-calendar"></span>
                            </span>
                        </div>
                    </div>
                    <div class="form-group">
                        <input type="text" class="form-control" name="location" id="location" placeholder="Location" />
                    </div>
                    <input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
                    <button type="submit" class="btn btn-default">Add Event</button>
                </form>
            </div>
        </div>
        <div class="table-responsive">
            <table id="calendarTable" class="table table-striped table-bordered">
                <thead>
                    <tr>
                        <th>Subject</th>
                        <th>Start</th>
                        <th>End</th>
                        <th>Location</th>
                        <th>Response</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach (var calendarEvent in Model)
                    {
                        <tr>
                            <td>
                                @{
                                    RouteValueDictionary idVal = new RouteValueDictionary();
                                    idVal.Add("eventId", calendarEvent.Id);
                                    @Html.ActionLink(calendarEvent.Subject, "Detail", idVal);
                                }
                            </td>
                            <td>
                                @string.Format("{0} ({1})", calendarEvent.Start.DateTime, calendarEvent.Start.TimeZone)
                            </td>
                            <td>
                                @string.Format("{0} ({1})", calendarEvent.End.DateTime, calendarEvent.End.TimeZone)
                            </td>
                            <td>
                                @{
                                    if (null != calendarEvent.Location)
                                    {
                                        @calendarEvent.Location.DisplayName
                                    }
                                }
                            </td>
                            <td>
                                @{
                                    if (null != calendarEvent.ResponseStatus.Response)
                                    {
                                          @calendarEvent.ResponseStatus.Response.Value
                                    }
                                }
                            </td>
                        </tr>
                    }
                </tbody>
            </table>
        </div>
        <div class="btn btn-group-sm">
            @{
                Dictionary<string, object> attributes = new Dictionary<string, object>();
                attributes.Add("class", "btn btn-default");

                if (null != ViewBag.NextLink)
                {
                    RouteValueDictionary routeValues = new RouteValueDictionary();
                    routeValues.Add("nextLink", ViewBag.NextLink);
                    @Html.ActionLink("Next Page", "Index", "Calendar", routeValues, attributes);
                }
            }
        </div>

    </div>
</div>
  ```
  
1. Create a new **View** to get event details and either Accept or TentativelyAccept or Decline it.
  1. Expand the **Views** folder in **MailCalendarWebApp**. Right-click **Calendar** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Detail.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Calendar/Detail.cshtml** with the following:
  
  ```asp
@model Microsoft.Graph.Event
@{ ViewBag.Title = "Detail"; }
<h2>@Model.Subject</h2>
@section scripts {
    <script type="text/javascript">
$(function () {
    $('#start-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
    $('#end-picker').datetimepicker({ format: 'YYYY-MM-DDTHH:mm:ss', sideBySide: true });
});
    </script>
}
<div class="row" style="margin-top:50px;">
    <div class="col-sm-12">
        @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
        {
            <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
        }
        <div class="panel panel-default">
            <div class="panel-body">
                <table>
                    <tbody>
                        <tr>
                            <form class="form-inline" action="/Calendar/Accept" method="post">
                                <input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
                                <button type="submit" name="Accept" class="btn btn-default">Accept</button>
                            </form>
                            <form class="form-inline" action="/Calendar/Tentative" method="post">
                                <input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
                                <button type="submit" name="Tentative" class="btn btn-default">Tentative</button>
                            </form>
                            <form class="form-inline" action="/Calendar/Decline" method="post">
                                <input type="hidden" name="eventId" value="@Request.Params["eventId"]" />
                                <button type="submit" name="Decline" class="btn btn-default">Decline</button>
                            </form>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        <div class="table-responsive">
            <table id="calendarTable" class="table table-striped table-bordered">
                <tbody>
                    <tr>
                        <td>Organizer:</td>
                        <td>
                            @Model.Organizer.EmailAddress.Name
                        </td>
                    </tr>
                    <tr>
                        <td>Start:</td>
                        <td>
                            @string.Format("{0} ({1})", Model.Start.DateTime, Model.Start.TimeZone)
                        </td>
                    </tr>
                    <tr>
                        <td>End:</td>
                        <td>
                            @string.Format("{0} ({1})", Model.End.DateTime, Model.End.TimeZone)
                        </td>
                    </tr>
                    <tr>
                        <td>Location:</td>
                        <td>
                            @{
                                if (null != Model.Location)
                                {
                                    @Model.Location.DisplayName
                                }
                            }
                        </td>
                    </tr>
                    <tr>
                        <td>Response:</td>
                        <td>
                            @{
                                if (null != Model.ResponseStatus.Response)
                                {
                                    @Model.ResponseStatus.Response.Value
                                }
                            }
                        </td>
                    </tr>
                        <td>Body:</td>
                        <td>
                            <pre>
                                @{
                                    @Model.Body.Content
                                }
                            </pre>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        <div class="btn btn-group-sm">
            @{
                Dictionary<string, object> attributes = new Dictionary<string, object>();
                attributes.Add("class", "btn btn-default");

                if (null != ViewBag.NextLink)
                {
                    RouteValueDictionary routeValues = new RouteValueDictionary();
                    routeValues.Add("nextLink", ViewBag.NextLink);
                    @Html.ActionLink("Next Page", "Index", "Calendar", routeValues, attributes);
                }
            }
        </div>

    </div>
</div>
  ```

### Run the app

1. Press **F5** to begin debugging.
1. When prompted, login with your Office 365 administrator account.
1. Click the **Outlook Mail API** or **Outlook Calendar API** link in the navigation bar at the top of the page.
1. Try out the app!

Congratulations! In this exercise you have created an MVC application that uses Microsoft Graph to view and manage Mails and Calendar in your mailbox.