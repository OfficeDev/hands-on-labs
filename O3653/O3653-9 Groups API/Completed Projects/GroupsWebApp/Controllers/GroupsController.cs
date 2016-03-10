using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.Graph;
using GroupsWebApp.Auth;
using GroupsWebApp.TokenStorage;
using Newtonsoft.Json;
using System.IO;

namespace GroupsWebApp.Controllers
{
  public class GroupsController : Controller
  {
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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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
          TempData["error"] = ex.Error.Message;
        }
      }

      return RedirectToAction("Index");
    }

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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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

        // Adding by email is a two-step process

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
          TempData["error"] = ex.Error.Message;
        }
      }

      return RedirectToAction("Members", new { groupId = groupId });
    }

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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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
          TempData["error"] = ex.Error.Message;
        }
      }

      return RedirectToAction("Conversations", new { groupId = groupId });
    }

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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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
          TempData["error"] = ex.Error.Message;
        }
      }

      return RedirectToAction("Calendar", new { groupId = groupId });
    }

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
        return RedirectToAction("Index", "Error", new { message = ex.Error.Message });
      }
    }

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
          // TEMP WORKAROUND
          if (!ex.Error.Message.Equals("An unexpected error occurred during deserialization."))
          {
            TempData["error"] = ex.Error.Message;
          }
        }
      }

      return RedirectToAction("Files", new { groupId = groupId });
    }

    private GraphServiceClient GetGraphServiceClient()
    {
      string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
      SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

      string tenantId = System.Security.Claims.ClaimsPrincipal.Current
          .FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

      string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");

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
      }),
      // WORKAROUND 
      new HttpProvider(new Serializer(new JsonSerializerSettings { TypeNameHandling = TypeNameHandling.None })));

      return client;
    }
  }
}