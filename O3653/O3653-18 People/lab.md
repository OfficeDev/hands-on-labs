#Microsoft Graph People API

In this lab, you will use Microsoft Graph to integrate the Office 365 People API with an ASP.NET MVC application.

## Get an Office 365 developer environment
To complete the exercises below, you will require an Office 365 developer environment. Navigate to https://tryoffice.azurewebsites.net and use the code `BuildChallenge` to get an administrator username and password to one. 

## Prerequisites
  1. You must have the OData v4 Client Code Generator add-in installed. In Visual Studio, go to **Tools** > **Extensions and Updates**, select "Online" from the left-most treeview, then search for "Odata v4 Client Code Generator", and click install.

## Exercise 1: Create a new project using Azure Active Directory v1 authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v1 Starter Project** template, register a new application
in the developer portal, and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **New** **Project**.
  1. Search the installed templates for **Graph** and select the **Graph AAD Auth v1 Starter Project** template.
  2. Name the new project **PeopleGraphWeb** and click **OK**.
  3. Note that for this quickstart, the AppId and AppSecret are already filled in in the web.config.

2. Press F5 to compile and launch your new application in the default browser.
   1. Once the Graph and AAD v2 Auth Endpoint Starter page appears, click **Sign in** and log in to your Office 365 account.
   2. Review the permissions that the application is requesting, and click **Accept**.
   3. Now that you are signed into your application, exercise 1 is complete!

## Exercise 2: Add a reference to the Graph API beta namespace

1. Right-click the project and select **add item**.
   1. Select **Visual C#** > **Code** > **Odata Client**.
   2. Name the file Graph.tt and click **Add**.
2. Edit the Graph.tt file.
   1. Edit MetadataDocumentUri to be "https://graph.microsoft.com/beta/$metadata".
   2. Edit NamespacePrefix to be "PeopleGraphWeb.Service".
3. Build the project.


## Exercise 3: Add the people controller and call the People API.
1. Right-click the **Controllers** folder and select **Add...** > **Controller...**. 
   1. Select **MVC5 Controller - Empty** and click **Add**.
   2. Name the controller **PeopleController** and click **Add**.
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
  
             string tenantId = System.Security.Claims.ClaimsPrincipal.Current
                 .FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
  
              string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");
  
             AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"],
               ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
  
             return await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
         }
  
         public async Task<Service.GraphService> GetService()
          {
             string token = await GetToken();
              Service.GraphService service = new Service.GraphService(new Uri("https://graph.microsoft.com/beta/"));
              service.BuildingRequest += (sender, e) => e.Headers.Add("Authorization", "Bearer " + token);
              return service;
          }
  
         private IEnumerable<T> Search<T>(Microsoft.OData.Client.DataServiceContext dataServiceContext,   Microsoft.OData.Client.DataServiceQuery<T> path, string searchString)
         {
              return dataServiceContext.Execute<T>(new Uri(path.RequestUri, "?$search=\"" + searchString + "\""));
          }
  ```
  
  
4. Add the index action that will list the relevant people for the logged-in user.
  
  ```c#
          [Authorize]
          public async Task<ActionResult> Index()
          {
              var service = await GetService();
              return View(service.Me.People.ToList());
          }
  ```

5. Add the view for the index controller. 
   1. Right-click the views folder and click **add** > **new folder**.
   2. Rename the folder to People.
   3. Right-click People and select **Add** > **View...**.
   4. Name the view "Index" and select Template "Empty".
   5. Set the contents of the file to the following:
  
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
      
  @foreach (var item in Model.People) {
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
1. Locate the **Views/Shared** folder in the project.
2. Open the **_Layout.cshtml** file found in the **Views/Shared** folder.
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
   
   
   2. Update that navigation to replace the "Graph API" link with "People"
       and connect this to the controller you just created.
       
       

   ```asp
   <ul class="nav navbar-nav">
       <li>@Html.ActionLink("Home", "Index", "Home")</li>
       <li>@Html.ActionLink("About", "About", "Home")</li>
       <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
       <li>@Html.ActionLink("People", "Index", "People")</li>
   </ul>
   ```
   

### Verify that the project works
1. In **Visual Studio**, hit **F5** to begin debugging.
2. When prompted, log in with your Office 365 Account.
3. Click the link **People** on the top of the home page.
4. Verify that your application displays the top relevant people for the current logged-in user.

## Exercise 4: Add support for people search and the details page
1. Add the search and details controllers:
  ```c#
        [Authorize]
        public async Task<ActionResult> Search(string searchText, string topic)
        {
            var service = await GetService();
            var searchString = string.IsNullOrWhiteSpace(topic) ? searchText : searchText + " topic:" + topic;
            return View("Index", Search(service, service.Me.People, searchString));
        }
  
        [Authorize]
        public async Task<ActionResult> Details(string id)
        {
            var service = await GetService();
            return View(service.Me.People.ByKey(id).GetValue());
        }
  ```
  
2. Add the details view and update the index view to support search.
   1. Right-click on the People folder under views and select **Add** > **View…**. 
      1. Set **View Name** to "Details".
      2. Set **Template** to "Details".
      3. Set **Model Class** to "Person (PersonGraphWeb.Service)".
   2. Edit the **Views/People/Index.cshtml**.
      1. Locate the table element and add the following code right above it:
  
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
  
  This will allow the user to enter search strings that will be passed to the search controller.

3. Verify the search and details features work.
  1. In **Visual Studio**, hit **F5** to begin debugging.
  2. When prompted, log in with your Office 365 Account.
  3. Click the link **People** on the top of the home page.
  4. Verify that your application displays the top relevant people for the current logged-in user.
  5. Click **details** to and verify the details are shown.
  6. Go back to the index and enter a search term into the search field, then click **Search**.
    For example: 
      * Search with the text: "Dennis Dehin" and see the fuzzy matched result Denis Dehenne is returned.
      * Search with the text: "Azis Hasoneh" and see the fuzzy matched result Aziz Hassouneh is returned.
      * Search with the topic: XT2000

## Exercise 5: Add support for working with

1. Add the following method to the people controller:
  ```c#
        [Authorize]
        public async Task<ActionResult> RelatedPeople(string id)
        {
            var service = await GetService();
            return View("Index", service.User.ByKey(id).People);
        }
  ```
    
  Notice the code re-uses the index view to display the results so another view is not needed.
  
2. Edit the People index view and add a new column to the table that links to the related people action:
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
            @Html.DisplayName("Display Name");
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
   1. In **Visual Studio**, hit **F5** to begin debugging.
   2. When prompted, log in with your Office 365 Account.
   3. Click the link **People** on the top of the home page.
   4. Selected **RelatedPeople** for a user and verify the related contacts are shown.

***
Congratulations, dedicated quick start developer! In this exercise, you have created an application that uses Microsoft Graph People API. This quick start ends here.  But don't stop here - there's plenty more to explore with the Microsoft Graph.

## Next Steps and Additional Resources:  
- See this training and more on http://dev.office.com/
- Learn about and connect to the Microsoft Graph at https://graph.microsoft.io
