# Microsoft Graph for OneDrive for Business
In this lab, you will use Microsoft Graph to integrate OneDrive for Business
with an ASP.NET MVC application.

## Prerequisites
1. You must have an Office 365 tenant and Microsoft Azure subscription to
   complete this lab. If you do not have one, the lab for **O3651-7 Setting up
   your Developer environment in Office 365** shows you how to obtain a trial.
2. You must have Visual Studio 2015 with Update 1 installed.
3. You must have the Graph AAD Auth v2 Started Project template installed.

## Exercise 1: Create a new project using Azure Active Directory v2 authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v2 Start Project** template, register a new application
in the developer portal, and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **New**, **Project**.
   1. Search the installed templates for **Graph** and select the
      **Graph AAD Auth v2 Starter Project** template.
   2. Name the new project **GraphFilesWeb** and click **OK**.
   3. Open the **Web.config** file and find the **appSettings** element. This is where you will need to add your appId and app secret you will generate in the next step.
2. Launch the [Application Registration Portal](https://apps.dev.microsoft.com)
   to register a new application.
      1. Sign into the portal using your Office 365 username and password.
      2. Click **Add an App** and type **Graph Files Quick Start** for the application name.
      3. Copy the **Application Id** and paste it into the value for **ida:AppId** in your project's **web.config** file.
      3. Under **Application Secrets** click **Generate New Password** to create a new client secret for your app.
      4. Copy the displayed app password and paste it into the value for **ida:AppSecret** in your project's **web.config** file.
      5. Modify the **ida:AppScopes** value to include the required **https://graph.microsoft.com/files.readwrite** scope.

```xml
<configuration>
  <appSettings>
    <!-- ... -->
    <add key="ida:AppId" value="paste application id here" />
    <add key="ida:AppSecret" value="paste application password here" />
    <!-- ... -->
    <!-- Specify scopes in this value. Multiple values should be comma separated. -->
    <add key="ida:AppScopes" value="https://graph.microsoft.com/user.read,https://graph.microsoft.com/files.readwrite" />
  </appSettings>
  <!-- ... -->
</configuration>
```
3. Add a redirect URL to enable testing on your localhost.
   1. Right click on **GraphFilesWeb** and click on **Properties** to open the project properties.
   2. Click on **Web** in the left navigation.
   3. Copy the **Project Url** value.
   4. Back on the Application Registration Portal page, click **Add Platform** and then **Web**.
   5. Paste the value of **Project Url** into the **Redirect URIs** field.
   6. Scroll to the bottom of the page and click **Save**.

4. Press F5 to compile and launch your new application in the default browser.
   1. Once the Graph and AAD v2 Auth Endpoint Starter page appears, click **Sign in** and login to your Office 365 account.
   2. Review the permissions the application is requesting, and click **Accept**.
   3. Now that you are signed into your application, exercise 1 is complete!

## Exercise 2: Access OneDrive for Business files through Microsoft Graph SDK

In this exercise, you will build on exercise 1 to connect to the Microsoft Graph
SDK can perform CRUD operations associated with the files in OneDrive for Business
or OneDrive.

1. Add a reference to the Microsoft Graph SDK to your project
  1. In the **Solution Explorer** right click on the **GraphFilesWeb** project and select **Manage NuGet Packages...**.
  2. Click **Browse** and search for **Microsoft.Graph**.
  3. Select the Microsoft Graph SDK and click **Install**.

2. Create a new controller to process the requests for files and send them to Graph API.
  1. Find the **Controllers** folder under **GraphFilesWeb**, right click on it and select **Add** then **Controller**.
  2. Select **MVC 5 Controller - Empty** and click **Add**.
  3. Change the name of the controller to **FilesController** and click **Add**.

3. **Add** the following reference to the top of the `FilesController` class.

```csharp
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.Graph;
using GraphFilesWeb.Auth;
using GraphFilesWeb.TokenStorage;
```

4. Add the following code to the `FilesController` class to initialize a new
   **GraphServiceClient** and generate an access token for the Graph API:

```csharp
private GraphServiceClient GetGraphServiceClient()
{
    string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
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

5. Add the following code to the `FilesController` class to return the view of items
   in the root of the user's OneDrive.

```csharp
[Authorize]
public async Task<ActionResult> Index(int? pageSize)
{
    var client = GetGraphServiceClient();

    // Define and store the current page size
    pageSize = pageSize ?? 10;
    ViewBag.PageSize = pageSize.Value;

    var request = client.Me.Drive.Root.Children.Request().Top(pageSize.Value);

    var results = await request.GetAsync();

    return View(results);
}
```

6. Add the following code to the `FilesController` to delete a file from the user's OneDrive:

```csharp
[Authorize]
public async Task<ActionResult> Delete(string itemId, string etag)
{
    var client = GetGraphServiceClient();

    // Build a request and set the If-Match header with the etag
    var request = client.Me.Drive.Items[itemId].Request(new List<Option> { new HeaderOption("If-Match", etag) });

    // Submit the delete request
    await request.DeleteAsync();

    return Redirect("/Files");
}
```

7. Add the following code to the `FilesController` to upload a new file to the user's OneDrive:

```csharp
[Authorize]
public async Task<ActionResult> Upload()
{
    var client = GetGraphServiceClient();

    foreach (string key in Request.Files)
    {
        var fileInRequest = Request.Files[key];
        if (fileInRequest != null && fileInRequest.ContentLength > 0)
        {
            var filename = System.IO.Path.GetFileName(fileInRequest.FileName);
            var request = client.Me.Drive.Root.Children[filename].Content.Request();
            var uploadedFile = await request.PutAsync<DriveItem>(fileInRequest.InputStream);
        }
    }

    return Redirect("/Files");
}
```


### Step 3. Code the MVC Application
Now you will code the MVC application to allow navigating the OneDrive file
collection using the Microsoft Graph.

1. Locate the **Views/Shared** folder in the project.
2. Open the **_Layout.cshtml** file found in the **Views/Shared** folder.
    1. Locate the part of the file that includes a few links at the top of the
       page... it should look similar to the following code:

```asp
  <div class="navbar-collapse collapse">
      <ul class="nav navbar-nav">
          <li><a asp-controller="Home" asp-action="Index">Home</a></li>
          <li><a asp-controller="Home" asp-action="About">About</a></li>
          <li><a asp-controller="Home" asp-action="Contact">Contact</a></li>
      </ul>
      @await Html.PartialAsync("_LoginPartial")
  </div>
```

    2. Update that navigation to have a new link (the **Files (Graph)** link
       added below) as well as a reference to the login control you just created:

```asp
 <div class="navbar-collapse collapse">
     <ul class="nav navbar-nav">
         <li><a asp-controller="Home" asp-action="Index">Home</a></li>
         <li><a asp-controller="Home" asp-action="About">About</a></li>
         <li><a asp-controller="Home" asp-action="Contact">Contact</a></li>
         <li><a asp-controller="Files" asp-action="Index">OneDrive Files</a></li>
     </ul>
     @await Html.PartialAsync("_LoginPartial")
 </div>
```

3. Right-click the **Controllers** folder and select **Add/New Item** and then
   select **MVC Controller Class**.
   1. In the **Add New Item** dialog select **MVC Controller Class** and name
      the file **FilesController.cs**.

4. **Add** the following references to the top of the file.

```c#
  using Microsoft.AspNet.Authorization;
```

4. **Replace** the **Index** method with the following code to read files.

```csharp
  [Authorize]
  public async Task<ActionResult> Index(int? pageSize)
  {
      string accessToken = null;
      FileRepository repository = new FileRepository(accessToken);

      // setup paging defaults if not provided
      pageSize = pageSize ?? 10;

      // setup paging for the IU
      ViewBag.PageSize = pageSize.Value;

      var results = await repository.GetMyFilesAsync(pageSize.Value);
      return View(results);
  }
```
5. Create a new **View** for the FilesController:
   1. Right click on the **Views** folder in **GraphFilesWeb** and select **Add** then **New Folder**.
   2. Name the folder **Files**.
   3. Right click on the new folder **Files** and select **Add** then **New Item**.
   4. Select **MVC View Page** and leave the filename **Index.cshtml** and click **Add**.

6. **Replace** all of the code in the **Files/Index.cshtml** with the following:

```asp
  @model IEnumerable<Microsoft.Graph.DriveItem>

  @{ ViewBag.Title = "My Files"; }

  <h2>My Files</h2>

  <div class="row" style="margin-top:50px;">
      <div class="col-sm-12">
          <div class="table-responsive">
              <table id="filesTable" class="table table-striped table-bordered">
                  <thead>
                      <tr>
                          <th></th>
                          <th>ID</th>
                          <th>Name</th>
                          <th>Created</th>
                          <th>Modified</th>
                      </tr>
                  </thead>
                  <tbody>
                      @foreach (var file in Model)
                      {
                          <tr>
                              <td>
                                  @{
  //Place delete control here
                                  }
                              </td>
                              <td>
                                  @file.Id
                              </td>
                              <td>
                                  <a href="@file.WebUrl">@file.Name</a>
                              </td>
                              <td>
                                  @file.CreatedDateTime
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
                  //Place Paging controls here
              }
          </div>
      </div>
  </div>
```

7. In **Visual Studio**, hit **F5** to begin debugging.
8. When prompted, log in with your **Organizational Account**.
9. Click the link **OneDrive Files** on the top of the home page.
10. Verify that your application displays files from the user's OneDrive.

  ![](Images/08.png)

11. Stop debugging.
12. In the **FilesController.cs** file, **add** the following code to delete files.

```csharp
  [Authorize]
  public async Task<ActionResult> Delete(string id, string etag)
  {
      FileRepository repository = new FileRepository(accessToken: null);
      if (id != null)
      {
          await repository.DeleteItemAsync(id, etag);
      }

      return Redirect("/Files");
  }
```

13. In the **Index.cshtml** file under **Views/Files** folder, **add** the
    following code under the comment `Place delete control here`.

```csharp
    Dictionary<string, object> attributes1 = new Dictionary<string, object>();
    attributes1.Add("class", "btn btn-warning");

    RouteValueDictionary routeValues1 = new RouteValueDictionary();
    routeValues1.Add("name", file.Id);
    routeValues1.Add("etag", file.eTag);
    @Html.ActionLink("X", "Delete", "Files", routeValues1, attributes1);
```

14. Press **F5** to begin debugging.

15. Test the delete functionality in the application.

Congratulations! In this exercise you have created an MVC application that uses
Microsoft Graph to return and manage files in a OneDrive!
