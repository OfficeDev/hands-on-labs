# Integrate with Files in OneDrive for Business
## Exercise 1: Create an ASP.NET MVC5 Application that authenticates with Office 365
### Step 1: Create your ASP.NET MVC Project
In this exercise, you will create the ASP.NET MVC5 application and register it with Azure Active Directory.
Launch Visual Studio 2017 as an administrator.
In Visual Studio select File/New/Project.
In the New Project dialog, select Templates/Visual C#/Web and click ASP.NET Web Application. Name the new project OneDriveWeb and then click OK.
 ![Screenshot 1](filegraphimg/image003.png)
Note: Make sure you enter the exact same name for the Visual Studio Project that is specified in these lab instructions (OneDriveWeb). The Visual Studio Project name becomes part of the namespace in the code. The code inside these instructions depends on the namespace matching the Visual Studio Project name specified in these instructions. If you use a different project name the code will not compile unless you adjust all the namespaces to match the Visual Studio Project name you enter when you create the project.
In the New ASP.NET Project dialog, click MVC and then click Change Authentication.
Select Work And School Accounts.  Add the domain of your test tenancy, which is typically of the form LODSC123456.onmicrosoft.com to the Domain: text box. Check Read directory data and click OK.
![Screenshot 2](filegraphimg/image005.png)
Once the New ASP.NET Project dialog appears like the following screenshot, click OK.
![Screenshot 3](filegraphimg/image007.png)
Congratulations... at this point your app is configured with Azure AD and leverages OpenID Connect and OWIN to facilitate the authentication process!
### Step 2. Grant App Necessary Permissions
Browse to the Azure Management Portal (portal.azure.com) and sign in with your temporary Office 365 Organizational Account.
In the left-hand navigation list, click Azure Active Directory. In the middle of the portal, then click 
"App Registrations."
Select the directory you share with your Office 365 subscription.
Search for the app with the ida:ClientId value that was created in the web.config file in exercise 1.
![Screenshot 4](filegraphimg/image009.png)
Select the application.
Click the Required Permissions button.
Click the + Add button.
Under 1 Select an API, select the Microsoft Graph option and click the Select Button.
For the new Microsoft Graph application permission entry, under the Delegated Permissions area, select the following permission:
Have full access to all files user can access
Click the Select button at the bottom of the page.
Click the Done button at the bottom of the “1 Select an API 2 Select Permissions” pane.
At this point you can test the authentication flow for your application.
In Visual Studio, press F5. The browser will automatically launch taking you to the HTTPS start page for the web application.
To sign in, click the Sign In link in the upper-right corner.
Login using the temporary Office 365 Organizational Account admin you have received.  This will typically be the of the form “admin@LODSC123456.onmicrosoft.com”
Upon a successful login, since this will be the first time you have logged into this app, Azure AD will present you with the common consent dialog that looks similar to the following image:
![Screenshot 5](filegraphimg/image010.jpg)
Click Accept to approve the app's permission request on your data in Office 365.
You will then be redirected back to your web application. However notice in the upper right corner, it now shows your email address and the Sign Out link.
In Visual Studio, press Shift+F5 to stop debugging.
## Exercise 2: Use Microsoft Graph for OneDrive for Business
### Step 1.  Add Common Operations
In this exercise, you will create a repository object for wrapping CRUD operations associated with files in OneDrive for Business.
In the Solution Explorer, create a new folder named Util.
Right-click the Util folder and select Add/Class, in the Add New Item dialog, name the new class SettingsHelper and click Add to create the new source file for the class.
At the top of the SettingsHelper.cs file, remove all the using statements and add the following using statements.
```
using System.Configuration;
```
Implement the new class SettingsHelper using the following class definition.
```
public class SettingsHelper
{ 
    public static string ClientId
    {
        get { return ConfigurationManager.AppSettings["ida:ClientId"]; }
    } 
    public static string ClientSecret
    {
        get { return ConfigurationManager.AppSettings["ida:ClientSecret"]; }
    }
    public static string AzureAdInstance
    {
        get { return ConfigurationManager.AppSettings["ida:AADInstance"]; }
    }
    public static string AzureAdTenantId
    {
        get { return ConfigurationManager.AppSettings["ida:TenantId"]; }
    }
    public static string GraphResourceUrl
    {
        get { return "https://graph.microsoft.com/v1.0/"; }
    }
    public static string AzureAdGraphResourceURL
    {
        get { return "https://graph.microsoft.com/"; }
    }
    public static string AzureAdAuthority
    {
        get { return AzureAdInstance + AzureAdTenantId; }
    }
    public static string ClaimTypeObjectIdentifier
    {
        get { return "http://schemas.microsoft.com/identity/claims/objectidentifier"; }
    }
}
```
Assembly references are not added to the shared projects in ASP.NET MVC, rather they are added to the actual client projects. 
Therefore you need to add the following NuGet packages manually.
Open the Package Manager Console: View/Other Windows/Package Manager Console.
Enter each line below in the console, one at a time, pressing ENTER after each one. NuGet will install the package and all dependent packages:
```
PM> Install-Package Microsoft.Graph
```
In the Solution Explorer, locate the Models folder in the OneDriveWeb project.
Right-click the Models folder and select Add/Class.
In the Add New Item dialog, name the new class FileRepository.cs.
![Screenshot 6](filegraphimg/image013.png)
Click Add.
At the top of the FileRepository.cs file, remove all the using statements and add the following using statements.
``` 
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using OneDriveWeb.Util;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
```
Add a function named GetGraphAccessTokenAsync to the FileRepository class to retrieve an Access Token.
``` 
private async Task<string> GetGraphAccessTokenAsync()
{
    var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
    var userObjectId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;
    var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret);
    var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);
    AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.AzureAdAuthority, new ADALTokenCache(signInUserId));
    var result = await authContext.AcquireTokenSilentAsync(SettingsHelper.AzureAdGraphResourceURL, clientCredential, userIdentifier);
    return result.AccessToken;
} 
```
Add a function named GetGraphServiceAsync to the FileRepository class.
```
private async Task<GraphServiceClient> GetGraphServiceAsync()
{
    var accessToken = await GetGraphAccessTokenAsync();
    var graphserviceClient = new GraphServiceClient(SettingsHelper.GraphResourceUrl,
                                new DelegateAuthenticationProvider(
                                    (requestMessage) =>
                                    {
                                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                                        return Task.FromResult(0);
                                    }));
    return graphserviceClient;
}
```
Add the following method to get a list of all the items (folders and files) within the root of the user's OneDrive:
```
public async Task<List<DriveItem>> GetMyFiles(int pageIndex, int pageSize)
{
    try
    {
        var graphServiceClient = await GetGraphServiceAsync();
        var requestFiles = await graphServiceClient.Me.Drive.Root.Children.Request().GetAsync(); 
        return requestFiles.CurrentPage.OrderBy(i => i.Name).Skip(pageIndex * pageSize).Take(pageSize).ToList();
    }
    catch
    {
        throw;
    }
}
```
Add the following method to the FileRepository class to delete a single file from the user's OneDrive for Business drive: 
```
public async Task DeleteFile(string id)
{
    try
    {
        var graphServiceClient = await GetGraphServiceAsync();
        await graphServiceClient.Me.Drive.Items[id].Request().DeleteAsync();
    }
    catch
    {
        throw;
    }
}
```
Add the following method to the FileRepository class to upload a single file to the user's OneDrive for Business:
```
public async Task<DriveItem> UploadFile(System.IO.Stream filestream, string filename)
{
    try
    {
        var graphServiceClient = await GetGraphServiceAsync();
        var newItem = await graphServiceClient.Me.Drive.Root.Children.Request()
            .AddAsync(
            new DriveItem
            {
                Name = filename,
                File = new File()
            });
        return await graphServiceClient.Me.Drive.Items[newItem.Id].Content.Request().PutAsync<DriveItem>(filestream);
    }
    catch
    {
        throw;
    }
}
```
## Step 2. Add views to the MVC Application
Now you will code the MVC application to allow navigating the OneDrive for Business file collection using the Microsoft Graph.
Locate the Views/Shared folder in the project.
Open the _Layout.cshtml file found in the Views/Shared folder
Locate the part of the file that includes a few links at the top of the page... it should look similar to the following code:
``` 
<div class="navbar-collapse collapse">
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
    </ul>
    @Html.Partial("_LoginPartial")
</div>
``` 
 Update that navigation to have a new link (the Files (Graph) link added below) as well as a reference to the login control you just created:
```
<div class="navbar-collapse collapse">
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Files (Graph)", "Index", "Files")</li>
    </ul>
    @Html.Partial("_LoginPartial")
</div>
```
Right-click the Controllers folder and select Add/Controller.
In the Add Scaffold dialog, select MVC 5 Controller - Empty and click Add.
In the Add Controller dialog, give the controller the name FilesController and click Add.
At the top of the file, remove all the using statements and add the following using statements.
``` 
using OneDriveWeb.Models;
using System.Threading.Tasks;
using System.Web.Mvc;
```
Replace the Index method with the following code to read files.
```
[Authorize]
public async Task<ActionResult> Index(int? pageIndex)
{
    FileRepository repository = new FileRepository();
    const int pageSize = 10;
    if (pageIndex == null)
    {
     pageIndex = 1;
    }
    var files = await repository.GetMyFiles((int)pageIndex - 1, pageSize);
    ViewBag.pageIndex = pageIndex;
    ViewBag.morePagesAvailable = files.Count < pageSize ? false : true;
    return View(files);
}
```
Within the FilesController class, right click the View() at the end of the Index() method and select Add View.
Within the Add View dialog, set the following values:
View Name: Index.
Template: Empty (without model).
> Leave all other fields blank & unchecked.
Click Add.
Replace all of the code in the file with the following:
```
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
                                                      <th>Title</th>
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
                                                                       @file.CreatedDateTime.Value.UtcDateTime.ToLocalTime()
                                                               </td>
                                                               <td>
                                                                  @file.LastModifiedDateTime.Value.UtcDateTime.ToLocalTime()
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
In Visual Studio, hit F5 to begin debugging.
When prompted, log in with your Organizational Account.
Click the link Files (Graph) on the top of the home page.
Verify that your application displays files from the OneDrive for Business library.
Stop debugging.
## Step 3. Add Additional Code to Handle Upload and Download
In the FilesController.cs file, add the following code to upload and delete files.
```
  [Authorize]
  public async Task<ActionResult> Upload()
  { 
      FileRepository repository = new FileRepository();
      foreach (string key in Request.Files)
      {
          if (Request.Files[key] != null && Request.Files[key].ContentLength > 0)
          {
              var file = await repository.UploadFile(
                  Request.Files[key].InputStream,
                  Request.Files[key].FileName.Split('\\')[Request.Files[key].FileName.Split('\\').Length - 1]);
          }
      }
     return Redirect("/Files");
  }
  [Authorize]
  public async Task<ActionResult> Delete(string name)
  {
      FileRepository repository = new FileRepository();
      if (name != null)
      {
          await repository.DeleteFile(name);
      }
      return Redirect("/Files");
  }
```
In the Index.cshtml file under Views/Files folder, add the following code under the comment Place delete control here.
```
  Dictionary<string, object> attributes1 = new Dictionary<string, object>();
  attributes1.Add("class", "btn btn-warning");
  RouteValueDictionary routeValues1 = new RouteValueDictionary();
  routeValues1.Add("name", file.Id);
  routeValues1.Add("etag", file.ETag);
  @Html.ActionLink("X", "Delete", "Files", routeValues1, attributes1);
```
Add the following code under the comment Place Paging controls here:
```
  var pageLinkAttributes = new Dictionary<string, object> { { "class", "btn btn-default" } };
  int pageIndex = ViewBag.pageIndex;
  // do prev link if not on first page
  if (pageIndex > 1)
  {
      var routeValues = new RouteValueDictionary { { "pageIndex", pageIndex - 1 } };
      @Html.ActionLink("Previous Page", "Index", "Files", routeValues, pageLinkAttributes);
  }
  // do next link if current page = max page size
  if (ViewBag.morePagesAvailable)
  {
      var routeValues = new RouteValueDictionary { { "pageIndex", pageIndex + 1 } };
          @Html.ActionLink("Next Page", "Index", "Files", routeValues, pageLinkAttributes);
  }
```
Add the following code to the bottom of the Index.cshtml file to create an upload control.
```
  <div class="row" style="margin-top:50px;">
      <div class="col-sm-12">
          @using (Html.BeginForm("Upload", "Files", FormMethod.Post, new { enctype = "multipart/form-data" }))
          {
              <input type="file" id="file" name="file" class="btn btn-default" />
              <input type="submit" id="submit" name="submit" value="Upload" class="btn btn-default" />
          }
      </div>
  </div>
```
Press F5 to begin debugging.
Test the paging, upload, and delete functionality in the application.
![Screenshot 7](filegraphimg/image014.png)
Congratulations! In this exercise you have created an MVC application that uses Microsoft Graph to to return and manage files in a OneDrive for Business file collection.
 