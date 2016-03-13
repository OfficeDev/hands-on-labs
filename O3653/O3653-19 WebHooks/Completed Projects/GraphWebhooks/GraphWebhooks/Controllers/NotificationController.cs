using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using GraphWebhooks.Auth;
using GraphWebhooks.Models;
using GraphWebhooks.SignalR;
using GraphWebhooks.TokenStorage;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.AccessControl;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace GraphWebhooks.Controllers
{
    public class NotificationController : Controller
    {
        public ActionResult LoadView()
        {
            return View("Notification");
        }

        // The notificationUrl endpoint that's registered with the webhooks subscription.
        [HttpPost]
        public async Task<ActionResult> Listen()
        {

            // Validate the new subscription by sending the token back to MS Graph.
            // This response is required for each subscription.
            if (Request.QueryString["validationToken"] != null)
            {
                var token = Request.QueryString["validationToken"];
                return Content(token, "plain/text");
            }

            // Parse the received notifications.
            else
            {
                try
                {
                    var notifications = new Dictionary<string, Notification>();
                    using (var inputStream = new System.IO.StreamReader(Request.InputStream))
                    {
                        JObject jsonObject = JObject.Parse(inputStream.ReadToEnd());
                        if (jsonObject != null)
                        {

                            // Notifications are sent in a 'value' array.
                            JArray value = JArray.Parse(jsonObject["value"].ToString());
                            foreach (var notification in value)
                            {
                                Notification current = JsonConvert.DeserializeObject<Notification>(notification.ToString());
                                current.ResourceData = JsonConvert.DeserializeObject<ResourceData>(notification["resourceData"].ToString());

                                var subscriptionParams = (Tuple<string, string>)HttpRuntime.Cache.Get("subscriptionId_" + current.SubscriptionId);
                                if (subscriptionParams != null)
                                {
                                    // Verify the message is from Microsoft Graph.
                                    if (current.ClientState == subscriptionParams.Item1)
                                    {
                                        // Just keep the latest notification for each resource.
                                        // No point pulling data more than once.
                                        notifications[current.Resource] = current;
                                    }
                                }
                            }
                            if (notifications.Count > 0)
                            {

                                // Query for the changed messages. 
                                await GetChangedMessagesAsync(notifications.Values);
                            }
                        }
                    }
                    return new HttpStatusCodeResult(200);
                }
                catch (Exception)
                {

                    // TODO: Handle the exception.
                    // Return a 200 so the service doesn't resend the notification.
                    return new HttpStatusCodeResult(200);
                }
            }
        }

        // Get information about the changed messages and send to browser via SignalR
        // A production application would typically queue a background job for reliability.
        public async Task GetChangedMessagesAsync(IEnumerable<Notification> notifications)
        {
            List<Message> messages = new List<Message>();
            string serviceRootUrl = "https://graph.microsoft.com/v1.0/";

            // Get an access token and add it to the client.
            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common", "");
            AuthenticationContext authContext = new AuthenticationContext(authority);
            ClientCredential credential = new ClientCredential(ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"]);



            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            foreach (var notification in notifications)
            {
                var subscriptionParams = (Tuple<string, string>)HttpRuntime.Cache.Get("subscriptionId_" + notification.SubscriptionId);
                string refreshToken = subscriptionParams.Item2;
                AuthenticationResult authResult = await authContext.AcquireTokenByRefreshTokenAsync(refreshToken, credential, "https://graph.microsoft.com");

                // Send the 'GET' request.
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, serviceRootUrl + notification.Resource);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(continueOnCapturedContext: false);

                // Get the messages from the JSON response.
                if (response.IsSuccessStatusCode)
                {
                    string stringResult = await response.Content.ReadAsStringAsync();
                    var type = notification.ResourceData.ODataType;
                    if (type == "#Microsoft.Graph.Message")
                    {
                        messages.Add(JsonConvert.DeserializeObject<Message>(stringResult));
                    }
                }
            }
            if (messages.Count > 0)
            {
                NotificationService notificationService = new NotificationService();
                notificationService.SendNotificationToClient(messages);
            }
        }
    }
}