using System.Threading.Tasks;
using System.Web.Mvc;
using System.Security.Claims;

using InboxSync.Models;
using InboxSync.Helpers;
using InboxSync.Auth;

using System.IO;
using Newtonsoft.Json;
using Microsoft.AspNet.SignalR;
using InboxSync.Hubs;

using PagedList;

namespace InboxSync.Controllers
{
  public class MailController : Controller
  {
    // GET: Mail
    [System.Web.Mvc.Authorize]
    public ActionResult Index()
    {
      return RedirectToAction("Inbox");
    }

    // GET: Mail/Inbox
    [System.Web.Mvc.Authorize]
    public async Task<ActionResult> Inbox(int? page)
    {
      // The user's database ID record
      string userId = (string)Session["user_id"];

      ViewBag.UserId = userId;
      ViewBag.IsSubscribed = await UserManager.IsUserSubscribed(userId);

      int pageSize = 20;
      int pageNumber = (page ?? 1);

      // Get the total count of messages for the user (for calculating number of pages)
      long totalCount = await UserManager.GetUsersMessageCount(userId);

      // Get the current page of messages
      var messages = await UserManager.GetUsersMessages(userId, pageSize, pageNumber);

      // Return a PagedList to the view
      return View(new StaticPagedList<Message>(messages, pageNumber, pageSize, (int)totalCount));
    }

    // POST Mail/SyncInbox
    [HttpPost]
    [System.Web.Mvc.Authorize]
    public async Task<ActionResult> SyncInbox()
    {
      if (null == Session["user_id"])
      {
        return Redirect("/");
      }

      string userId = (string)Session["user_id"];

      try
      {
        await UserManager.SyncUsersInbox(userId);
      }
      catch (System.Exception ex)
      {
        return RedirectToAction("Index", "Error", new { message = ex.Message });
      }

      return RedirectToAction("Inbox");
    }

    // POST Mail/ResetSyncState
    [HttpPost]
    [System.Web.Mvc.Authorize]
    public async Task<ActionResult> ResetSyncState()
    {
      if (null == Session["user_id"])
      {
        return Redirect("/");
      }

      string userId = (string)Session["user_id"];
      await UserManager.ResetSyncState(userId);

      return RedirectToAction("Inbox");
    }

    // POST Mail/Subscribe
    [HttpPost]
    [System.Web.Mvc.Authorize]
    public async Task<ActionResult> Subscribe()
    {
      if (null == Session["user_id"])
      {
        return Redirect("/");
      }

      string userId = (string)Session["user_id"];

      string notificationUrl = Url.Action("Notify", "Mail", null, Request.Url.Scheme);
      await UserManager.SubscribeForInboxUpdates(userId, notificationUrl);

      return RedirectToAction("Inbox");
    }

    // POST Mail/Subscribe
    [HttpPost]
    [System.Web.Mvc.Authorize]
    public async Task<ActionResult> Unsubscribe()
    {
      if (null == Session["user_id"])
      {
        return Redirect("/");
      }

      string userId = (string)Session["user_id"];

      await UserManager.UnsubscribeForInboxUpdates(userId);

      return RedirectToAction("Inbox");
    }

    // POST Mail/Notify
    [HttpPost]
    public async Task<ActionResult> Notify(string validationToken)
    {
      // Check if this is a validation request.
      if (!string.IsNullOrEmpty(validationToken))
      {
        // To validate return the validationToken back in the response
        return Content(validationToken, "text/plain");
      }

      // If it's not a validation request, then this is a notification.

      Stream postBody = Request.InputStream;
      postBody.Seek(0, SeekOrigin.Begin);
      string notificationJson = new StreamReader(postBody).ReadToEnd();

      NotificationPayload notification = JsonConvert.DeserializeObject<NotificationPayload>(notificationJson);

      string userId = await UserManager.UpdateInboxBySubscription(notification.Notifications[0].SubscriptionId);

      if (!string.IsNullOrEmpty(userId))
      {
        var hub = GlobalHost.ConnectionManager.GetHubContext<NotificationHub>();
        string connId = await UserManager.GetUserSignalRConnection(userId);
        hub.Clients.Client(connId).refreshInboxPage();
      }

      return Content("");
    }
  }
}