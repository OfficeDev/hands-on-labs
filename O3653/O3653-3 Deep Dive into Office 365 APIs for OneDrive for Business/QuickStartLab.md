# Microsoft Graph for OneDrive for Business
In this lab, you will use Microsoft Graph to integrate OneDrive for Business
with an ASP.NET MVC5 application.

## Prerequisites
1. You must have an Office 365 tenant and Microsoft Azure subscription to
   complete this lab. If you do not have one, the lab for **O3651-7 Setting up
   your Developer environment in Office 365** shows you how to obtain a trial.
2. You must have Visual Studio 2015 with Update 1 installed.


## Step 1: Use Azure Active Directory v2 end point to create an access token

* For apps targeting Microsoft accounts and work or school accounts, follow
  quick start **O3653-XXX Using Azure Active Directory v2 end point with ASP.NET MVC5**
* For apps targeting work or school accounts only, follow quick start
  **O3653-YYY Using Azure Active Directory Authentication Library with ASP.NET MVC5**

The remainder of this quick start challenge assumes that you have followed one
of these quick starts and can generate an OAuth **access_token** to make calls
to the Microsoft Graph API.

## Step 2: Access OneDrive for Business content from an ASP.NET MVC5 application

In this exercise, you will use the Microsoft Graph SDK to perform CRUD operations
associated with files in OneDrive for Business.

1. In the **Solution Explorer**, locate the **project.json** file under
   **GraphFilesWeb** and open it.
   1. Find the **dependencies** section and add a reference to the Microsoft
      Graph SDK package: `"Microsoft.Graph": "1.0.0"`
   2. Find **frameworks** and remove **dnxcore50** from the list. The Microsoft
      Graph SDK is not compatible with DNX Core 5.0.
   3. After saving the file, Visual Studio will automatically restore the package
      and make it available in the project.
2. Right-click the **GraphFilesWeb** project and select **Add/Class**.
3. In the **Add New Item** dialog, name the new class **FileRepository.cs**.
4. Click **Add**.

  ![](Images/07.png)

5. **Add** the following reference to the top of the `FileRepository` class.

```csharp
  using Microsoft.Graph;
```

6. Add the following code to the `FileRepository` class to initialize a new
   **GraphServiceClient** with the access token obtained from Step 1.

```csharp
private readonly string graphAccessToken;
private readonly GraphServiceClient graphClient;

public FileRepository(string accessToken)
{
    graphAccessToken = accessToken;
    graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (System.Net.Http.HttpRequestMessage request) =>
    {
        request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + graphAccessToken);
    }));
}
```

7. Add the following methods to get a list of all the items (folders and files)
   within the root of the user's OneDrive:

```csharp
public async Task<IChildrenCollectionPage> GetMyFilesAsync(int pageSize)
{
    // build the request for items in the root folder
    var request = graphClient.Me.Drive.Root.Children.Request().Top(pageSize);

    // get there results of the request
    var results = await request.GetAsync();
    return results;
}
```

8. Add the following method to the `FileRepository` class to delete a single
   file from the user's OneDrive:

```csharp
public async Task<bool> DeleteItemAsync(string id, string etag)
{
    // create request to delete the item
    var request = graphClient.Me.Drive.Items[id].Request();

    // Execute the delete action on this request
    try
    {
        await request.DeleteAsync();
        return true;
    }
    catch (Exception)
    {
        return false;
    }
}
```

9. Add the following method to the `FileRepository` class to upload a single
   file to the user's OneDrive:

```c#
public async Task<DriveItem> UploadFileAsync(System.IO.Stream filestream, string filename)
{
    // Create a request to upload the file using simple PUT action
    var request = graphClient.Me.Drive.Root.ItemWithPath(filename).Content.Request();

    // Submit the request with the contents of the filestream
    var result = await request.PutAsync<DriveItem>(filestream);
    return result;
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
