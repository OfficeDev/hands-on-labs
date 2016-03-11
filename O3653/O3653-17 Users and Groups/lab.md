### TODO
1. Nuget package installation directions needed? Currently there is no mention of Nuget packages in this doc. If package isn't installed, **GraphServiceClient** will not be defined. (Dan) I think you need to find the nuget using Manage Nuget Packages. Or we can put the nuget on the machine and have VS configured with the correct package source location.

# Lab 17: Microsoft Graph People Picker

## What you'll learn
In this lab, you will create an ASP.NET MVC application that uses the Microsoft Graph client SDK to create a people picker. It will search for users in your tenant's directory and show their profile information, including their picture.

## Prerequisites
1. Visual Studio 2015 with Update 1
2. The Graph AAD Auth v1 Started Project template installed
3. An administrator account for an Office 365 tenant. This is required because you'll be using the client credentials of an Azure application that's configured to request admin-level permissions.

## Step 1: Create a new project using Graph AAD Auth v1 Started Project template
In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v1 Start Project** template.

1. Open Visual Studio 2015 and select **File/New/Project**.
2. Search the installed templates for **Graph** and select the
      **Graph AAD Auth v1 Starter Project** template. This starter project template scaffolds some auth infrastructure for you, so that you can focus on calling the Microsoft Graph.
3. Name the new project **PeoplePicker** and click **OK**.
> NOTE: Make sure you use the exact same name that is specified in these instructions for your Visual Studio project. Otherwise, your namespace name will differ from the one in these instructions and your code will not compile.
    ![](images/VSProject.JPG)
4. Your new application is ready to go! *Required*: hit F5 to to restore the NuGet packages required by the project. This will compile and launch your new application in the default browser.  You can sign in to the app using the O365 tenant administrator account provided to you.

## Step 2: Implement user search bar using the Microsoft Graph SDK

2. Create a new controller: Right click on the **Controller** folder and select **Add**, **Controller**. Select **MVC 5 Controller - Empty**, click **Add** and then name the new controller **UserSearchController**.
3. Create an associated view by right clicking the function **Index()**, **Add View**, and click **Add**. The view is created at **Views\UserSearch\Index.cshtml**.
4. In **UserSearchController.cs**, replace the auto-generated **using** directives with
    ```c#
    using System;
    using System.IO;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Web.Mvc;
    using System.Threading.Tasks;
    using System.Security.Claims;
    using Microsoft.Graph;
    using PeoplePicker.TokenStorage;
    using PeoplePicker.Auth;
    ```
4. In **UserSearchController.cs**, insert code into the **UserSearchController** class to intialize a **GraphServiceClient**, later used to make our calls to Microsoft Graph. The **GraphServiceClient** is initialized by obtaining an access token through the `GetUserAccessToken` helper function.

    ```c#
    public class UserSearchController : Controller
    {
        public static string appId = ConfigurationManager.AppSettings["ida:AppId"];
        public static string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        public static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];

        private GraphServiceClient GetGraphServiceClient()
        {
            string userObjId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = string.Format(aadInstance, tenantID, "");
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

            // Create an authHelper using the the app Id and secret and the token cache
            AuthHelper authHelper = new AuthHelper(authority,appId,appSecret,tokenCache);

            // Request an accessToken and provide the original redirect URL from sign-in
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
            {
                string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
            }));

            return client;
        }
    }
    ```
5. In **UserSearchController.cs**, insert code into the **Index()** function to render the view the first time the page loads.

    ```csharp
    public ActionResult Index()
    {
        // Return an empty list of people when the page first loads
        List<User> people = new List<User>();
        return View(people);
    }
    ```

6. In the same file, add the following function to make the Graph query for users. The view passes a **searchString** and this function builds the query and returns the results to the view.

    ```csharp
    [HttpPost]
    public async Task<ActionResult> Index(FormCollection fc, string searchString)
    {
        // Search for users with name or mail that includes searchString
        var client = GetGraphServiceClient();

        List<User> people = new List<User>();

        // Graph query for users, filtering by displayName, givenName, surname, UPN, mail, and mailNickname
        // Only query for displayName, userPrincipalName, id of matching users through select
        try {
            var result = await client.Users.Request().Top(7).Filter("startswith(displayName,'" + searchString +
            "') or startswith(givenName,'" + searchString +
            "') or startswith(surname,'" + searchString +
            "') or startswith(userPrincipalName,'" + searchString +
            "') or startswith(mail,'" + searchString +
            "') or startswith(mailNickname,'" + searchString + "')").Select("displayName,userPrincipalName,id").GetAsync();

            // Add users to the list and return to the view
            foreach (User u in result) {
                people.Add(u);
            }
        }
        catch(Exception)
        {
            return View("Error");
        }
        return View(people);
    }
    ```

6. Replace the contents of **Views\UserSearch\Index.cshtml** with the following code. This renders a search bar and a table to display the results.

    ```xml
    @using Microsoft.Graph
    @model List<User>
    @{
        ViewBag.Title = "User search";
    }
    <h2>@ViewBag.Title</h2>

    @using (Html.BeginForm())
    {
        <div class="input-group">
            <input type="text" name="searchString" class="form-control" placeholder="Search by name or email...">
            <span class="input-group-btn">
                <input type="submit" value="Search" class="btn btn-default" />
            </span>
        </div>
    }
    <table class="table table-bordered table-striped table-hover">
        @foreach (var user in Model)
        {
            <tr>
                <td>
                    @Html.DisplayFor(modelItem => user.DisplayName)
                </td>
                <td>
                    @Html.DisplayFor(modelItem => user.UserPrincipalName)
                </td>
            </tr>
        }

    </table>
    ```
7. Add the following code to **Views\Shared\\_Layout.cshtml**, directly under the line ``` <li>@Html.ActionLink("Contact", "Contact", "Home")</li>``` to add a link to the navbar at the top of the page.

    ```xml
    @if (Request.IsAuthenticated)
    {
    <li>@Html.ActionLink("User Search", "Index", "UserSearch")</li>
    }
    ```

8. Hit F5 to build your app! First sign in, then explore the navbar **User search** link. Search for users by complete or incomplete name or email. Try typing *a* and hitting the search button :)

## Step 3: Create the detailed user profile page
In this step, we'll enable selecting a user from the user search results table, for whom we'll show more details.
1. Right click the folder **Models**, **Add**, **Class**, and name it **Profile.cs**. Replace the contents of this class with the following.

    ```csharp
    using Microsoft.Graph;
    using System;

    namespace PeoplePicker.Models
    {
        public class Profile
        {
            public String photo { get; set; }
            public User user { get; set; }
        }
    }
    ```

2. Add an **onclick** event to the table row element in **Views\UserSearch\Index.cshtml** to navigate to a different page when the table row is clicked.
    ```xml
    <tr onclick="location.href = '@(Url.Action("ShowProfile", "UserSearch", new { userId = user.Id }))'">
    ```
3. Add the following directive in **UserSearchController.cs** to use the model we just created:
    ```csharp
    using PeoplePicker.Models;
    ```
4. Add the following function to **UserSearchController.cs**. This function is passed a userId from the user selected in the table, and builds the query to get the full profile information of the user and their photo. This is passed to a new view through a Profile object.
    ```csharp
    public async Task<ActionResult> ShowProfile(string userId)
    {
        // Show the profile of a user after a user is clicked from the search
        var client = GetGraphServiceClient();
        Profile profile = new Profile();

        try {
            // Graph query for user details by userId
            profile.user = await client.Users[userId].Request().GetAsync();
            profile.photo = "";

            // Graph query for user photo by userId
            var photo = await client.Users[userId].Photo.Content.Request().GetAsync();

            if (photo != null)
            {
                // Convert to MemoryStream for ease of rendering
                using (MemoryStream stream = (MemoryStream)photo)
                {
                    string toBase64Photo = Convert.ToBase64String(stream.ToArray());
                    profile.photo = "data:image/jpeg;base64, " + toBase64Photo;
                }
            }
        }
        catch (Exception)
        {
            return View("Error");
        }

        return View(profile);
    }
    ```
4. Create a new view under **Views\UserSearch** and name it **ShowProfile.cshtml**. Replace the contents of this file with the following.
    ```xml
    @model PeoplePicker.Models.Profile
    @{
        ViewBag.Title = "Profile";
    }

    <h2>Profile</h2>

    @{ if (Model.photo != "")
        {
            <div style="height: 200px">
                <img src="@Model.photo" alt="Photo" border="0" style="width:auto;max-height: 100%" />
            </div>
        }
    }

    @{if (Model.user.GivenName != null && Model.user.Surname != null)
        {
            <h3> @Model.user.GivenName @Model.user.Surname </h3>
        }
    }

    <table class="table table-bordered table-striped">
        <tr>
            <td>Display name</td>
            <td>
                @if (!string.IsNullOrEmpty(Model.user.DisplayName)){
                    @Html.DisplayFor(modelItem => Model.user.DisplayName)
                }
            </td>
        </tr>
        <tr>
            <td>Mail</td>
            <td>
                @if (!string.IsNullOrEmpty(Model.user.Mail))
                {
                    @Html.DisplayFor(modelItem => Model.user.Mail)
                }
            </td>
        </tr>
        <tr>
            <td>Job title</td>
            <td>
                @if (!string.IsNullOrEmpty(Model.user.JobTitle))
                {
                    @Html.DisplayFor(modelItem => Model.user.JobTitle)
                }
            </td>
        </tr>
        <tr>
            <td>Mobile phone</td>
            <td>
                @if (!string.IsNullOrEmpty(Model.user.MobilePhone))
                {
                    @Html.DisplayFor(modelItem => Model.user.MobilePhone)
                }
            </td>
        </tr>
        <tr>
            <td>Office</td>
            <td>
                @if (!string.IsNullOrEmpty(Model.user.OfficeLocation))
                {
                    @Html.DisplayFor(modelItem => Model.user.OfficeLocation)
                }
            </td>
        </tr>
    </table>
    ```

5. Hit F5 to compile and try out the new ShowProfile page. Search for users and click on a user to see their details.

***
Hooray! Congratulations on creating your PeoplePicker app! You have created an MVC application that uses Microsoft Graph to search and view users in your tenant.
