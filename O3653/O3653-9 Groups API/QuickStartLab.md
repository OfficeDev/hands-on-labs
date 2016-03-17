# Connect to groups, add a member, see the group files and conversations
In this lab, you will use Microsoft Graph to integrate Groups, its files, conversations and events
with an ASP.NET MVC application.

## Exercise 1: Create a new project using Azure Active Directory authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v1 Starter Project** template and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **New**, **Project**.
  1. Search the installed templates for **Graph** and select the
    **Graph AAD Auth v1 Starter Project** template.
  1. Name the new project **GroupsWebApp** and click **OK**.
   
1. Press F5 to compile and launch your new application in the default browser.
  1. Once the Graph and AAD Auth Endpoint Starter page appears, click **Sign in** and login to your Office 365 adminsitrator account.
  1. Review the permissions the application is requesting, and click **Accept**.
  1. Now that you are signed into your application, exercise 1 is complete!
   
## Exercise 2: Access Groups through Microsoft Graph SDK

In this exercise, you will build on exercise 1 to connect to the Microsoft Graph
SDK and work with Office 365 Groups.

### Create the groups controller and use the Graph SDK

1. Add a reference to the Microsoft Graph SDK to your project
  1. In the **Solution Explorer** right click on the **GroupsWebApp** project and select **Manage NuGet Packages...**.
  1. Click **Browse** and search for **Microsoft.Graph**.
  1. Select the Microsoft Graph SDK and click **Install**.
  
1. Add a reference to the Bootstrap DateTime picker to your project
  1. In the **Solution Explorer** right click on the **GroupsWebApp** project and select **Manage NuGet Packages...**.
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
  1. Find the **Controllers** folder under **GroupsWebApp**, right click on it and select **Add** then **Controller**.
  1. Select **MVC 5 Controller - Empty** and click **Add**.
  1. Change the name of the controller to **GroupsController** and click **Add**.

1. **Add** the following reference to the top of the `GroupsController` class.

  ```csharp
  using System.Configuration;
  using System.Threading.Tasks;
  using Microsoft.Graph;
  using GroupsWebApp.Auth;
  using GroupsWebApp.TokenStorage;
  ```
  
1. Add the following code to the `GroupsController` class to initialize a new
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
  
### Work with groups
  
1. Add the following code to the `GroupsController` class to get all Office 365 groups.

  ```csharp
  // GET: Groups
  [Authorize]
  public async Task<ActionResult> Index(int? pageSize, string nextLink)
  {
    var client = GetGraphServiceClient();

    pageSize = pageSize ?? 25;

    // Filter to only return groups with the 'Unified' type,
    // which corresponds to Office 365 groups
    var request = client.Groups.Request().Top(pageSize.Value).Filter("groupTypes/any(c:c+eq+'Unified')");
    if (!string.IsNullOrEmpty(nextLink))
    {
      request = new GroupsCollectionRequest(nextLink, client, null);
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
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to display details about a group.

  ```csharp
  // GET: Groups/Detail?groupId=<id>
  [Authorize]
  public async Task<ActionResult> Detail(string groupId)
  {
    var client = GetGraphServiceClient();

    var request = client.Groups[groupId].Request();

    try
    {
      var result = await request.GetAsync();

      return View(result);
    }
    catch (ServiceException ex)
    {
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to retrieve the group's photo.

  ```csharp
  // GET: Groups/Photo?groupId=<id>
  [Authorize]
  public async Task<ActionResult> Photo(string groupId)
  {
    // This example retrieves the photo from the server every time.
    // In a real app, it would be better to cache the photo after the first
    // download and return from cache.
    var client = GetGraphServiceClient();

    var photoRequest = client.Groups[groupId].Photo.Content.Request();

    try
    {
      var photoStream = await photoRequest.GetAsync();

      return new FileStreamResult(photoStream, "image/jpeg");
    }
    catch (ServiceException ex)
    {
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to create a new group.

  ```csharp
  // POST: Groups/CreateGroup?groupName=<text>&groupAlias=<text>
  [Authorize]
  [HttpPost]
  public async Task<ActionResult> CreateGroup(string groupName, string groupAlias)
  {
    if (string.IsNullOrEmpty(groupName) || string.IsNullOrEmpty(groupAlias))
    {
      TempData["error"] = "Please enter a name and alias";
    }
    else
    {
      var client = GetGraphServiceClient();

      var request = client.Groups.Request();

      // Initialize a new group
      Group newGroup = new Group()
      {
        DisplayName = groupName,
        // The group's email will be set as groupAlias@<yourdomain>
        MailNickname = groupAlias,
        MailEnabled = true,
        SecurityEnabled = false,
        GroupTypes = new List<string>() { "Unified" }
      };

      try
      {
        Group createdGroup = await request.AddAsync(newGroup);
        return RedirectToAction("Detail", new { groupId = createdGroup.Id });
      }
      catch (ServiceException ex)
      {
        if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
        TempData["error"] = ex.Error.Message;
      }
    }

    return RedirectToAction("Index");
  }
  ```
  
### Work with group members
  
1. Add the following code to the `GroupsController` class to list a group's members.

  ```csharp
  // GET: Groups/Members?groupId=<id>
  [Authorize]
  public async Task<ActionResult> Members(string groupId, int? pageSize, string nextLink)
  {
    if (!string.IsNullOrEmpty((string)TempData["error"]))
    {
      ViewBag.ErrorMessage = (string)TempData["error"];
    }

    pageSize = pageSize ?? 25;

    var client = GetGraphServiceClient();

    var request = client.Groups[groupId].Members.Request().Top(pageSize.Value);
    if (!string.IsNullOrEmpty(nextLink))
    {
      request = new MembersCollectionWithReferencesRequest(nextLink, client, null);
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
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to add a member to a group.

  ```csharp
  // POST: Groups/AddMember?groupId=<id>&newMemberEmail=<email>
  [Authorize]
  [HttpPost]
  public async Task<ActionResult> AddMember(string groupId, string newMemberEmail)
  {
    if (string.IsNullOrEmpty(newMemberEmail))
    {
      TempData["error"] = "Please enter an email address";
    }
    else
    {
      var client = GetGraphServiceClient();

      // Adding by email address is a two-step process

      // First we need to get the user from Graph so we 
      // have the user's ID property
      var userRequest = client.Users[newMemberEmail].Request();

      // Then we pass the user entity to the member add request
      var request = client.Groups[groupId].Members.References.Request();

      try
      {
        var user = await userRequest.GetAsync();
        await request.AddAsync(user);
      }
      catch (ServiceException ex)
      {
        if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
        TempData["error"] = ex.Error.Message;
      }
    }

    return RedirectToAction("Members", new { groupId = groupId });
  }
  ```
  
### Work with group conversations
  
1. Add the following code to the `GroupsController` class to get a group's conversations.

  ```csharp
  // GET: Groups/Conversations?groupId=<id>
  [Authorize]
  public async Task<ActionResult> Conversations(string groupId, int? pageSize, string nextLink)
  {
    if (!string.IsNullOrEmpty((string)TempData["error"]))
    {
      ViewBag.ErrorMessage = (string)TempData["error"];
    }

    pageSize = pageSize ?? 25;

    var client = GetGraphServiceClient();

    var request = client.Groups[groupId].Conversations.Request().Top(pageSize.Value);
    if (!string.IsNullOrEmpty(nextLink))
    {
      request = new ConversationsCollectionRequest(nextLink, client, null);
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
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to add a new conversation.

  ```csharp
  // POST: Groups/AddConversation?groupId=<id>&topic=<text>&message=<text>
  [Authorize]
  [HttpPost]
  public async Task<ActionResult> AddConversation(string groupId, string topic, string message)
  {
    if (string.IsNullOrEmpty(topic) || string.IsNullOrEmpty(message))
    {
      TempData["error"] = "Please enter topic and message";
    }
    else
    {
      var client = GetGraphServiceClient();

      var request = client.Groups[groupId].Conversations.Request();

      // Build the conversation
      Conversation conversation = new Conversation()
      {
        Topic = topic,
        // Conversations have threads
        Threads = new ThreadsCollectionPage()
      };
      conversation.Threads.Add(new ConversationThread()
      {
        // Threads contain posts
        Posts = new PostsCollectionPage()
      });
      conversation.Threads[0].Posts.Add(new Post()
      {
        // Posts contain the actual content
        Body = new ItemBody() { Content = message, ContentType = BodyType.text }
      });

      try
      {
        await request.AddAsync(conversation);
      }
      catch (ServiceException ex)
      {
        if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
        TempData["error"] = ex.Error.Message;
      }
    }

    return RedirectToAction("Conversations", new { groupId = groupId });
  }
  ```

### Work with group calendars

1. Add the following code to the `GroupsController` class to retrieve the upcoming events on a group's calendar in the next 7 days.

  ```csharp
  // GET: Groups/Calendar?groupId=<id>
  [Authorize]
  public async Task<ActionResult> Calendar(string groupId, int? pageSize, string nextLink)
  {
    if (!string.IsNullOrEmpty((string)TempData["error"]))
    {
      ViewBag.ErrorMessage = (string)TempData["error"];
    }

    pageSize = pageSize ?? 25;

    var client = GetGraphServiceClient();

    // In order to use a calendar view, you must specify
    // a start and end time for the view. Here we'll specify
    // the next 7 days.
    DateTime start = DateTime.Today;
    DateTime end = start.AddDays(6);

    // These values go into query parameters in the request URL,
    // so add them as QueryOptions to the options passed ot the
    // request builder.
    List<Option> viewOptions = new List<Option>();
    viewOptions.Add(new QueryOption("startDateTime", 
      start.ToUniversalTime().ToString("s", System.Globalization.CultureInfo.InvariantCulture)));
    viewOptions.Add(new QueryOption("endDateTime", 
      end.ToUniversalTime().ToString("s", System.Globalization.CultureInfo.InvariantCulture)));

    var request = client.Groups[groupId].CalendarView.Request(viewOptions).Top(pageSize.Value);
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
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to create a new event on a group's calendar.

  ```csharp
  // POST Groups/AddEvent?groupId=<id>&subject=<text>&start=<text>&end=<text>&location=<text>
  [Authorize]
  [HttpPost]
  public async Task<ActionResult> AddEvent(string groupId, string subject, string start, string end, string location)
  {
    if (string.IsNullOrEmpty(subject) || string.IsNullOrEmpty(start) 
      || string.IsNullOrEmpty(end) || string.IsNullOrEmpty(location))
    {
      TempData["error"] = "Please fill in all fields";
    }
    else
    {
      var client = GetGraphServiceClient();

      var request = client.Groups[groupId].Events.Request();

      Event newEvent = new Event()
      {
        Subject = subject,
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
        if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
        TempData["error"] = ex.Error.Message;
      }
    }

    return RedirectToAction("Calendar", new { groupId = groupId });
  }
  ```
  
### Work with group files
  
1. Add the following code to the `GroupsController` class to get the list of files in the root of a group's OneDrive.

  ```csharp
  // GET: Groups/Files?groupId=<id>
  [Authorize]
  public async Task<ActionResult> Files(string groupId, int? pageSize, string nextLink)
  {
    if (!string.IsNullOrEmpty((string)TempData["error"]))
    {
      ViewBag.ErrorMessage = (string)TempData["error"];
    }

    pageSize = pageSize ?? 25;

    var client = GetGraphServiceClient();

    var request = client.Groups[groupId].Drive.Root.Children.Request().Top(pageSize.Value);
    if (!string.IsNullOrEmpty(nextLink))
    {
      request = new ChildrenCollectionRequest(nextLink, client, null);
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
      if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
      return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
    }
  }
  ```
  
1. Add the following code to the `GroupsController` class to create a new file in the root of a group's OneDrive.

  ```csharp
  // POST: Groups/AddFile?groupId=<id>
  [Authorize]
  [HttpPost]
  public async Task<ActionResult> AddFile(string groupId)
  {
    var selectedFile = Request.Files["file"];
    if (null == selectedFile || 0 == selectedFile.ContentLength)
    {
      TempData["error"] = "Please select a file to add";
    }
    else if (selectedFile.ContentLength > 4 * 1024 * 1024)
    {
      // Simple upload only supports files up to 4MB
      TempData["error"] = "Please select a file under 4 MB in size";
    }
    else
    {
      var client = GetGraphServiceClient();

      string fileName = Path.GetFileName(selectedFile.FileName);

      var request = client.Groups[groupId].Drive.Root.Children[fileName].Content.Request();

      try
      {
        var upload = await request.PutAsync<DriveItem>(selectedFile.InputStream);
      }
      catch (ServiceException ex)
      {
        if (ex.Error.Code == "InvalidAuthenticationToken") { return new EmptyResult(); }
        // TEMP WORKAROUND
        if (!ex.Error.Message.Equals("An unexpected error occurred during deserialization."))
        {
          TempData["error"] = ex.Error.Message;
        }
      }
    }

    return RedirectToAction("Files", new { groupId = groupId });
  }
  ```
  
### Create the groups view

In this section you'll wire up the Controller you created in the previous section
to an MVC view that will display the organization's groups.

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

  1. Update that navigation to replace the "Graph API" link with "Groups"
      and connect this to the controller you just created.

  ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Groups", "Index", "Groups")</li>
    </ul>
  ```
1. Create a new **View** for groups.
  1. Expand the **Views** folder in **GroupsWebApp**. Right-click **Groups** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Index.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Groups/Index.cshtml** with the following:
  
  ```asp
  @model IEnumerable<Microsoft.Graph.Group>

  @{ ViewBag.Title = "Groups"; }

  <h2>Groups</h2>

  <div class="row" style="margin-top:50px;">
      <div class="col-sm-12">
          @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
          {
          <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
          }
          <div class="panel panel-default">
              <div class="panel-body">
                  <form class="form-inline" action="/Groups/CreateGroup" method="post">
                      <div class="form-group">
                          <label for="groupName">Name</label>
                          <input type="text" class="form-control" name="groupName" id="groupName" placeholder="New Group" />
                      </div>
                      <div class="form-group">
                          <label for="groupAlias">Email alias</label>
                          <input type="text" class="form-control" name="groupAlias" id="groupAlias" placeholder="e.g. mygroup" />
                      </div>
                      <button type="submit" class="btn btn-default">Create Group</button>
                  </form>
              </div>
          </div>
          <div class="table-responsive">
              <table id="groupsTable" class="table table-striped table-bordered">
                  <thead>
                      <tr>
                          <th>Name</th>
                          <th>Description</th>
                          <th>Email address</th>
                      </tr>
                  </thead>
                  <tbody>
                      @foreach (var group in Model)
                      {
                      <tr>
                          <td>
                              @{
                                RouteValueDictionary idVal = new RouteValueDictionary();
                                idVal.Add("groupId", group.Id);
                                @Html.ActionLink(group.DisplayName, "Detail", idVal);
                              }
                          </td>
                          <td>
                              @group.Description
                          </td>
                          <td>
                              @group.Mail
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
                  @Html.ActionLink("Next Page", "Index", "Groups", routeValues, attributes);
                }
              }
          </div>
      </div>
  </div>
  ```
  
1. Create a new **View** for group detail.
  1. Expand the **Views** folder in **GroupsWebApp**. Right-click **Groups** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Detail.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Groups/Detail.cshtml** with the following:
  
  ```asp
  @model Microsoft.Graph.Group

  @{ ViewBag.Title = "Group Detail"; }

  <div class="panel panel-default">
      <div class="panel-body">
          <hr />
          <div class="media">
              <div class="media-left">
                  @{
                    RouteValueDictionary idVal = new RouteValueDictionary();
                    idVal.Add("groupId", Model.Id);
                    <img src="@Url.Action("Photo", idVal)" style="height: 128px; width: 128px;"/>
                  }
              </div>
              <div class="media-body">
                  <h4 class="media-heading">@Model.DisplayName</h4>
                  <p>@Model.Description</p>
              </div>
          </div>
          <hr />
          <ul class="list-group">
              <li class="list-group-item">@Html.ActionLink("View Members", "Members", idVal)</li>
              <li class="list-group-item">@Html.ActionLink("View Conversations", "Conversations", idVal)</li>
              <li class="list-group-item">@Html.ActionLink("View Calendar", "Calendar", idVal)</li>
              <li class="list-group-item">@Html.ActionLink("View Files", "Files", idVal)</li>
          </ul>
      </div>
  </div>
  ```
  
1. Create a new **View** for group members.
  1. Expand the **Views** folder in **GroupsWebApp**. Right-click **Groups** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Members.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Groups/Members.cshtml** with the following:
  
  ```asp
  @model IEnumerable<Microsoft.Graph.DirectoryObject>

  @{ ViewBag.Title = "Group Members"; }

  <h2>Group Members</h2>

  <div class="row" style="margin-top:50px;">
      <div class="col-sm-12">
          @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
          {
          <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
          }
          <div class="panel panel-default">
              <div class="panel-body">
                  <form class="form-inline" action="/Groups/AddMember" method="post">
                      <div class="form-group">
                          <label for="newMemberEmail">Email</label>
                          <input type="email" class="form-control" name="newMemberEmail" id="newMemberEmail" placeholder="user@contoso.com" />
                      </div>
                      <input type="hidden" name="groupId" value="@Request.Params["groupId"]" />
                      <button type="submit" class="btn btn-default">Add Member</button>
                  </form>
              </div>
          </div>
          <div class="table-responsive">
              <table id="membersTable" class="table table-striped table-bordered">
                  <thead>
                      <tr>
                          <th>Name</th>
                          <th>Email address</th>
                      </tr>
                  </thead>
                  <tbody>
                      @foreach (Microsoft.Graph.User user in Model)
                      { 
                      <tr>
                          <td>
                              @user.DisplayName
                          </td>
                          <td>
                              @user.Mail
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
                  @Html.ActionLink("Next Page", "Members", "Groups", routeValues, attributes);
                }
              }
          </div>
      </div>
  </div>
  ```
  
1. Create a new **View** for group conversations.
  1. Expand the **Views** folder in **GroupsWebApp**. Right-click **Groups** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Conversations.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Groups/Conversations.cshtml** with the following:
  
  ```asp
  @model IEnumerable<Microsoft.Graph.Conversation>

  @{ ViewBag.Title = "Conversations"; }

  <h2>Conversations</h2>

  <div class="row" style="margin-top:50px;">
      <div class="col-sm-12">
          @if (!string.IsNullOrEmpty(ViewBag.ErrorMessage))
          {
          <div class="alert alert-danger">@ViewBag.ErrorMessage</div>
          }
          <div class="panel panel-default">
              <div class="panel-body">
                  <form class="form-inline" action="/Groups/AddConversation" method="post">
                      <div class="form-group">
                          <label for="topic">Topic</label>
                          <input type="text" class="form-control" name="topic" id="topic" placeholder="Enter a topic for this conversation" />
                      </div>
                      <div class="form-group">
                          <label for="message">Message</label>
                          <input type="text" class="form-control" name="message" id="message" placeholder="Enter a message" />
                      </div>
                      <input type="hidden" name="groupId" value="@Request.Params["groupId"]" />
                      <button type="submit" class="btn btn-default">Add Conversation</button>
                  </form>
              </div>
          </div>
          <div class="table-responsive">
              <table id="conversationsTable" class="table table-striped table-bordered">
                  <thead>
                      <tr>
                          <th>Topic</th>
                          <th>Preview</th>
                      </tr>
                  </thead>
                  <tbody>
                      @foreach (var convo in Model)
                      {
                      <tr>
                          <td>
                              @(string.IsNullOrEmpty(convo.Topic) ? "(No subject)" : convo.Topic)
                          </td>
                          <td>
                              @convo.Preview
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
                  @Html.ActionLink("Next Page", "Conversations", "Groups", routeValues, attributes);
                }
              }
          </div>
      </div>
  </div>
  ```
  
1. Create a new **View** for group calendar.
  1. Expand the **Views** folder in **GroupsWebApp**. Right-click **Groups** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Calendar.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Groups/Calendar.cshtml** with the following:
  
  ```asp
  @model IEnumerable<Microsoft.Graph.Event>

  @{ ViewBag.Title = "Calendar"; }

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
                  <form class="form-inline" action="/Groups/AddEvent" method="post">
                      <div class="form-group">
                          <input type="text" class="form-control" name="subject" id="subject" placeholder="Subject" />
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
                      <input type="hidden" name="groupId" value="@Request.Params["groupId"]" />
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
                      </tr>
                  </thead>
                  <tbody>
                      @foreach (var calendarEvent in Model)
                      {
                      <tr>
                          <td>
                              @calendarEvent.Subject
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
                  @Html.ActionLink("Next Page", "Calendar", "Groups", routeValues, attributes);
                }
              }
          </div>
      </div>
  </div>
  ```
  
1. Create a new **View** for group files.
  1. Expand the **Views** folder in **GroupsWebApp**. Right-click **Groups** and select
      **Add** then **New Item**.
  1. Select **MVC View Page** and change the filename **Files.cshtml** and click **Add**.
  1. **Replace** all of the code in the **Groups/Files.cshtml** with the following:
  
  ```asp
  @model IEnumerable<Microsoft.Graph.DriveItem>

  @{ ViewBag.Title = "Files"; }

  <h2>Files</h2>

  @section scripts {
  <script type="text/javascript">
  $(function () {
      // Validate file size < 4 MB
      // We're using the "simple upload" method of
      // uploading files to OneDrive, which is limited
      // to 4MB.
      // See http://graph.microsoft.io/en-us/docs/api-reference/v1.0/api/item_uploadcontent
      $('#file-form').submit(function () {
          var fourMB = 4 * 1024 * 1024;
          var fileInput = $('#file');
          if (fileInput.get(0).files[0].size > fourMB) {
              alert('Maximum file size is 4 MB.');
              return false;
          }
      });
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
                  <form id="file-form" class="form-inline" action="/Groups/AddFile" method="post" enctype="multipart/form-data">
                      <div class="form-group">
                          <input type="file" size="50" accept=".txt" class="form-control btn btn-default" name="file" id="file" />
                      </div>
                      <input type="hidden" name="groupId" value="@Request.Params["groupId"]" />
                      <button type="submit" class="btn btn-default">Add File</button>
                  </form>
              </div>
          </div>
          <div class="table-responsive">
              <table id="filesTable" class="table table-striped table-bordered">
                  <thead>
                      <tr>
                          <th>Name</th>
                          <th>Created On</th>
                          <th>Last Modified By</th>
                          <th>Last Modified On</th>
                      </tr>
                  </thead>
                  <tbody>
                      @foreach (var file in Model)
                      {
                      <tr>
                          <td>
                              @file.Name
                          </td>
                          <td>
                              @file.CreatedDateTime
                          </td>
                          <td>
                              @file.LastModifiedBy.User.DisplayName
                          </td>
                          <td>
                              @file.LastModifiedDateTime
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
                  @Html.ActionLink("Next Page", "Files", "Groups", routeValues, attributes);
                }
              }
          </div>
      </div>
  </div>
  ```
  
### Run the app

1. Press **F5** to begin debugging.
1. When prompted, login with your Office 365 administrator account.
1. Click the **Groups** link in the navigation bar at the top of the page.
1. Try out the app!

Congratulations! In this exercise you have created an MVC application that uses Microsoft Graph to view and manage groups in Office 365!
