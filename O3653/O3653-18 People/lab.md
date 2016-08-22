#Microsoft Graph People API

In this lab, you will use the Microsoft Graph to integrate the Office 365 People API with an ASP.NET MVC application.

## Get an Office 365 developer environment
To complete the exercises below, you will require an Office 365 developer environment. Use the Office 365 tenant that you have been provided with for Microsoft Ignite.

## Prerequisites
  1. You must have the OData v4 Client Code Generator add-in installed. In Visual Studio, go to **Tools** > **Extensions and Updates**, select **Online** from the left-most treeview, and search for *OData v4 Client Code Generator*.

## Exercise 1: Create a new project using Azure Active Directory v2 authentication

In this first step, you will create a new ASP.NET MVC project using the **Graph AAD Auth v2 Starter Project** template, register a new application
in the developer portal, and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **File** > **New** > **Project**.
  1. Search the installed templates for **Graph** and select the **Graph AAD Auth v2 Starter Project** template.
  1. Name the new project **PeopleGraphWeb** and click **OK**.
  1. Open the **Web.config** file in the root directory and find the **appSettings** element. This is where you will add the app ID and app secret that you will generate in the next step.

2. Launch the Application Registration Portal by opening a browser to [apps.dev.microsoft.com](https://apps.dev.microsoft.com)
   to register a new application.
  1. Sign into the portal using your Office 365 username and password. The **Graph AAD Auth v2 Starter Project** template allows you to sign in with either a Microsoft account or an Office 365 for business account, but the "People" features work only with business and school accounts.
  1. Click **Add an app** and type **PeopleGraphQuickStart** for the application name.
  1. Copy the **Application Id** and paste it into the value for **ida:AppId** in your project's **Web.config** file.
  1. Under **Application Secrets** click **Generate New Password** to create a new client secret for your app.
  1. Copy the displayed app password and paste it into the value for **ida:AppSecret** in your project's **Web.config** file.
  1. Set the **ida:AppScopes** value to *User.ReadBasic.All,People.Read*.

  ```xml
  <configuration>
    <appSettings>
      <!-- ... -->
      <add key="ida:AppId" value="4b63ba37..." />
      <add key="ida:AppSecret" value="AthR0e75..." />
      <!-- ... -->
      <!-- Specify scopes in this value. Multiple values should be comma separated. -->
      <add key="ida:AppScopes" value="User.ReadBasic.All,People.Read" />
    </appSettings>
    <!-- ... -->
  </configuration>
  ```

3. Add a redirect URL to enable testing on your localhost.
  1. In Visual Studio, right-click **PeopleGraphWeb** > **Properties** to open the project properties.
  1. Click **Web** in the left navigation.
  1. Copy the **Project Url** value.
  1. Back on the Application Registration Portal page, click **Add Platform** > **Web**.
  1. Paste the project URL into the **Redirect URIs** field.
  1. At the bottom of the page, click **Save**.

4. Set the Start action to the **Account/Signout** action (to avoid a stale token error). 
  1. In Visual Studio, right-click **PeopleGraphWeb** > **Properties** to open the project properties.
  1. Click **Web** in the left navigation.
  1. Under **Start Action** choose **Specific Page** and enter *Account/SignOut*. 

5. Press F5 to compile and launch your new application in the default browser.
  1. Once the Graph and AAD v2 Auth Endpoint Starter page appears, and sign in with your Office 365 account.
  1. Review the permissions the application is requesting, and click **Accept**.
  1. Now that you are signed into your application, exercise 1 is complete!

## Exercise 2: Add a reference to the Graph API beta namespace

1. In Visual Studio, right-click **PeopleGraphWeb** and select **Add** > **New Item**.
   1. Select **Visual C#** > **Code** > **OData Client**.
   2. Name the file *Graph.tt* and click **Add**.

2. Edit the Graph.tt file.
   1. Set **MetadataDocumentUri** to `https://graph.microsoft.com/beta/$metadata`.
   2. Set **NamespacePrefix** to *PeopleGraphWeb.Service*.

3. Build the project.

## Exercise 3: Add the PeopleController and call the People API

1. Right-click the **Controllers** folder and select **Add > New Scaffolded Item**. 
   1. Select **MVC5 Controller - Empty** and click **Add**.
   2. Name the controller *PeopleController* and click **Add**.

2. Edit the using statements:
      
  ```C#
  using System;
  using System.Collections.Generic;
  using System.Configuration;
  using System.Linq;
  using System.Security.Claims;
  using System.Threading.Tasks;
  using System.Web.Mvc;
  using PeopleGraphWeb.Auth;
  using PeopleGraphWeb.TokenStorage;
  using PeopleGraphWeb.Service;
  ```
      
3. Add the following helper functions to the PeopleController class:
  
  ```C#
    public async Task<string> GetToken()
    {
        string userObjId = System.Security.Claims.ClaimsPrincipal.Current
            .FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
  
        SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);
  
        string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common", "");
  
        AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"],
            ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
  
        return await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
    }

    public Service.GraphService GetService(string token)
    {
        Service.GraphService service = new Service.GraphService(new Uri("https://graph.microsoft.com/beta/"));
        service.BuildingRequest += (sender, e) => e.Headers.Add("Authorization", "Bearer " + token);
        return service;
    }

    private IEnumerable<T> Search<T>(
        Microsoft.OData.Client.DataServiceContext dataServiceContext,
        Microsoft.OData.Client.DataServiceQuery<T> path,
        string searchString)
    {
        return dataServiceContext.Execute<T>(new Uri(path.RequestUri, "?$search=\"" + searchString + "\""));
    }
  ```
  
  
4. Edit the **Index** action to list the relevant people for the logged-in user:
  
  ```c#
    [Authorize]
    public async Task<ActionResult> Index()
    {
        var token = await GetToken();
        if (!string.IsNullOrEmpty(token))
        {
            var service = GetService(token);
            return View(service.Me.People.ToList());                
        }
        return RedirectToAction("SignOut", "Account");
    }
  ```

5. Add the view for the Index action. 
   1. In the **Views** folder, right-click the **People** folder and select **Add** > **View**.
   2. Set **View name** to *Index*.
   3. Set **Template** to *Empty (without model)* and click **Add**.
   3. Replace the contents of the file with the following:
  
  ```asp
  @model IEnumerable<PeopleGraphWeb.Service.Person>
  
  @{
      ViewBag.Title = "People";
  }

  <table class="table">
      <tr>
          <th>
              @Html.DisplayName("Display Name");
          </th>
          <th></th>
      </tr>
      
  @foreach (var item in Model) 
  {
      <tr>
         <td>
             @Html.DisplayFor(modelItem => item.DisplayName)
         </td>
          <td>
              @Html.ActionLink("Details", "Details", new { id=item.Id }) 
          </td>
      </tr>
  }

  </table>
    ```

### Edit the default layout to point to our new controller

1. In the **Views/Shared** folder, open the **_Layout.cshtml** file.
2. Locate the part of the file that includes a few links at the top of the page. It should look similar to the following code:
  
    ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("Graph API", "Graph", "Home")</li>
    </ul>
    ```

3.  Replace the "Graph API" link with a "People" link that connects to the controller you just created.
    
    ```asp
    <ul class="nav navbar-nav">
        <li>@Html.ActionLink("Home", "Index", "Home")</li>
        <li>@Html.ActionLink("About", "About", "Home")</li>
        <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
        <li>@Html.ActionLink("People", "Index", "People")</li>
    </ul>
    ```

### Verify that the project works

1. Press F5 to begin debugging.
2. Sign in with your Office 365 account and click the **People** link at the top of the page.
3. Verify that your application displays the top relevant people for the current logged-in user.

## Exercise 4: Add support for people search and the details page

1. In PeopleController, add the Search and Details actions:

  ```c#
    [Authorize]
    public async Task<ActionResult> Search(string searchText, string topic)
    {
        var token = await GetToken();
        if (!string.IsNullOrEmpty(token))
        {
            var service = GetService(token);
            var searchString = string.IsNullOrWhiteSpace(topic) ? searchText : searchText + " topic:" + topic;
            return View("Index", Search(service, service.Me.People, searchString));
        }
        return RedirectToAction("SignOut", "Account");
    }

    [Authorize]
    public async Task<ActionResult> Details(string id)
    {
        var token = await GetToken();
        if (!string.IsNullOrEmpty(token))
        {
            var service = GetService(token);
            return View(service.Me.People.ByKey(id).GetValue());
        }
        return RedirectToAction("SignOut", "Account");
    }
  ```
  
2. Add the Details view.
   1. Right-click **Views/People** folder and select **Add** > **View**. 
      1. Set **View name** to *Details*.
      2. Set **Template** to *Details*.
      3. Set **Model class** to *Person (PersonGraphWeb.Service)* and click **Add**.

3. Update the Index view to support search.
   1. In **Views/People/Index.cshtml**, locate the table element and add the following code right above it:
  
  ```asp
@{ using (Html.BeginForm("Search", "People", FormMethod.Get))
    {
        @Html.Label("Seach:")
        @Html.TextBox("searchText")
        @Html.Label("Topic:")
        @Html.TextBox("topic")
        <input type="submit" value="Search" />
    }
}
  ```
  
  This allows the user to enter search strings that will be passed to the search action.

4. Verify the search and details features work.
  1. Press F5 to begin debugging.
  2. Sign in with your Office 365 account and click the **People** link at the top of the page.
  3. Verify that your application displays the top relevant people for the current logged-in user.
  4. Click **Details** for a user and verify the details are shown.
  5. Go back to the index page and enter a search term into the **Search** field, then click **Search**.

    For example: 
      * Search with the text: *Dennis Dehin* and see the fuzzy matched result Denis Dehenne is returned.
      * Search with the text: *Azis Hasoneh* and see the fuzzy matched result Aziz Hassouneh is returned.
      * Search with the topic: *XT2000*

## Exercise 5: Add support for working with related people

1. Add the following method to the PeopleController class:
  ```c#
    [Authorize]
    public async Task<ActionResult> RelatedPeople(string id)
    {
        var token = await GetToken();
        if (!string.IsNullOrEmpty(token))
        {
            var service = GetService(token);
            return View("Index", service.Users.ByKey(id).People);               
        }
        return RedirectToAction("SignOut", "Account");
    }
  ```
    
  Notice the code re-uses the Index view to display the results so another view is not needed.
  
2. In **Views/People/Index.cshtml**, add a new column to the table that links to the related people action:
  ```asp
<td>
    @Html.ActionLink("Related People", "RelatedPeople", new { id=item.Id }) 
</td> 
  ```

  The table should now look like this:
  ```asp
<table class="table">
    <tr>
        <th>
            @Html.DisplayName("Display Name")
        </th>
        <th></th>
        <th></th>
    </tr>
  
    @foreach (var item in Model)
    {
        <tr>
            <td>
                @Html.DisplayFor(modelItem => item.DisplayName)
            </td>
            <td>
                @Html.ActionLink("Details", "Details", new { id = item.Id })
            </td>
            <td>
                @Html.ActionLink("Related People", "RelatedPeople", new { id = item.Id })
            </td> 
        </tr>
    }

</table>
    ```
      
3. Verify the search and details features work.
  1. Press F5 to begin debugging.
  2. Sign in with your Office 365 account and click the **People** link at the top of the page.
  3. Click **Related People** for a user and verify the related contacts are shown.

***
Congratulations, dedicated quick start developer! In this exercise, you have created an application that uses the Microsoft Graph People API. This quick start ends here. But don't stop here - there's plenty more to explore with the Microsoft Graph.

## Next Steps and Additional Resources:  
- See this training and more on http://dev.office.com/.
- Learn about the Microsoft Graph at https://graph.microsoft.io.
