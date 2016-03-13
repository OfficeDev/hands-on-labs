using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace InboxSync.Helpers
{
  public class OutlookHelper
  {
    // Used to set the base API endpoint, e.g. "https://outlook.office.com/api/beta"
    public string ApiEndpoint { get; set; }
    // Used to set the X-AnchorMailbox header, which helps to efficiently route
    // API requests to the correct server
    public string AnchorMailbox { get; set; }

    public OutlookHelper()
    {
      // Set default endpoint
      ApiEndpoint = "https://outlook.office.com/api/beta";
      AnchorMailbox = string.Empty;
    }

    // Used to make a REST API call to a URL
    public async Task<HttpResponseMessage> MakeApiCall(string method, string token, string apiUrl,
     string userEmail, string payload, string[] preferHeaders)
    {
      using (var httpClient = new HttpClient())
      {
        var request = new HttpRequestMessage(new HttpMethod(method), apiUrl);

        // Headers
        request.Headers.Authorization =
          new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        request.Headers.UserAgent.Add(
          new System.Net.Http.Headers.ProductInfoHeaderValue("dotnet-outlook-nosdk", "1.0"));
        request.Headers.Accept.Add(
          new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
        request.Headers.Add("return-client-request-id", "true");
        request.Headers.Add("X-AnchorMailbox", userEmail);

        if (preferHeaders != null)
        {
          foreach (string header in preferHeaders)
          {
            request.Headers.Add("Prefer", header);
          }
        }

        // Content
        if ((method.ToUpper() == "POST" || method.ToUpper() == "PATCH") &&
            !string.IsNullOrEmpty(payload))
        {
          request.Content = new StringContent(payload);
          request.Content.Headers.ContentType.MediaType = "application/json";
        }

        var apiResult = await httpClient.SendAsync(request);
        return apiResult;
      }
    }

    // Used to sync inbox messages
    public async Task<JObject> SyncInbox(string email, string accessToken, string deltaToken, string skipToken)
    {
      string syncEndpoint = this.ApiEndpoint + "/Me/MailFolders/Inbox/Messages";

      // Set up query parameters
      string query = "?$select=Subject,ReceivedDateTime,From,BodyPreview,IsRead&$orderby=ReceivedDateTime+desc";

      // Append sync state if present
      if (!string.IsNullOrEmpty(deltaToken))
      {
        // deltaToken is used to start a new sync after a prior sync completed
        query += "&$deltatoken=" + deltaToken;
      }
      else if (!string.IsNullOrEmpty(skipToken))
      {
        // skipToken is used during a sync when there are more results than the max page size
        query += "&$skiptoken=" + skipToken;
      }

      syncEndpoint += query;

      // Set the odata.track-changes to enable sync
      // Set the sync page size to 20 items
      string[] preferences = {
        "odata.track-changes",
        "odata.maxpagesize=20"
      };

      HttpResponseMessage syncResult = await this.MakeApiCall("GET", accessToken,
        syncEndpoint, email, null, preferences);

      string response = await syncResult.Content.ReadAsStringAsync();

      return JObject.Parse(response);
    }

    // Used to create a notification subscription on the user's inbox
    public async Task<JObject> CreateInboxSubscription(string email, string accessToken, string notificationUrl)
    {
      string subscribeEndpoint = this.ApiEndpoint + "/Me/Subscriptions";

      // Build the JSON payload
      var subscription = new JObject(
        new JProperty("@odata.type", "#Microsoft.OutlookServices.PushSubscription"),
        new JProperty("Resource", this.ApiEndpoint + "/Me/MailFolders/Inbox/Messages"),
        new JProperty("NotificationURL", notificationUrl),
        // We want to be notified if anything is new, changed, or deleted
        new JProperty("ChangeType", "Created, Updated, Deleted")
      );

      // POST the JSON to the /Subscriptions endpoint
      HttpResponseMessage subscribeResult = await this.MakeApiCall("POST", accessToken, subscribeEndpoint, email, subscription.ToString(), null);

      string response = await subscribeResult.Content.ReadAsStringAsync();

      return JObject.Parse(response);
    }

    // Used to delete a subscription
    public async Task DeleteSubscription(string email, string accessToken, string subscriptionId)
    {
      string unsubscribeEndpoint = this.ApiEndpoint + "/Me/Subscriptions/" + subscriptionId;

      HttpResponseMessage unsubscribeResult = await this.MakeApiCall("DELETE", accessToken, unsubscribeEndpoint, email, null, null);
    }
  }
}