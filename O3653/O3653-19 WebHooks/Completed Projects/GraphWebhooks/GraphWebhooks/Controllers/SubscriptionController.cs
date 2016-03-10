using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using GraphWebhooks.Auth;
using GraphWebhooks.Models;
using GraphWebhooks.TokenStorage;
using Newtonsoft.Json;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace GraphWebhooks.Controllers
{
    public class SubscriptionController : Controller
    {
        // GET: Subscription
        public ActionResult Index()
        {
            return View();
        }

        // Create webhooks subscriptions.
        [Authorize]
        public async Task<ActionResult> CreateSubscription()
        {

            // Get an access token and add it to the client.
            AuthenticationResult authResult = null;
            try
            {
                string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);
                string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
                string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");

                AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
                authResult = await authHelper.GetUserTokens(Url.Action("Index", "Home", null, Request.Url.Scheme));

            }
            catch (Exception ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Message, debug = ex.StackTrace });
            }

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Build the request.
            string subscriptionsEndpoint = "https://graph.microsoft.com/beta/subscriptions/";
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, subscriptionsEndpoint);
            var subscription = new Subscription
            {
                Resource =  "me/messages",
                ChangeType = "Created",
                NotificationUrl = ConfigurationManager.AppSettings["ida:NotificationUrl"],
                ClientState = Guid.NewGuid().ToString()
            }; 
   
            request.Content = new StringContent(JsonConvert.SerializeObject(subscription), System.Text.Encoding.UTF8, "application/json");

            // Send the request and parse the response.
            HttpResponseMessage response = await client.SendAsync(request);
            if (response.IsSuccessStatusCode)
            {

                // Parse the JSON response.
                string stringResult = await response.Content.ReadAsStringAsync();
                SubscriptionViewModel viewModel = new SubscriptionViewModel
                {
                    Subscription = JsonConvert.DeserializeObject<Subscription>(stringResult)
                };

                // This app temporarily stores the current subscription ID, refreshToken and client state. 
                // These are required so the NotificationController, which is not authenticated can retrieve an access token keyed from subscription id
                // Production apps typically use some method of persistent storage.
                HttpRuntime.Cache.Insert("subscriptionId_" + viewModel.Subscription.SubscriptionId, 
                    Tuple.Create(viewModel.Subscription.ClientState, authResult.RefreshToken), null, DateTime.MaxValue, new TimeSpan(24, 0, 0), System.Web.Caching.CacheItemPriority.NotRemovable, null);

                // Save the latest subscription ID, so we can delete it later and filter teh view on it.
                Session["SubscriptionId"] = viewModel.Subscription.SubscriptionId;
                return View("Subscription", viewModel);

            }
            else
            {
                return RedirectToAction("Index", "Error", new { message = response.StatusCode, debug = response.Content });
            }
        }

        // Delete the current webhooks subscription and sign the user out.
        [Authorize]
        public async Task<ActionResult> DeleteSubscription()
        {
            string subscriptionId = (string) Session["SubscriptionId"];

            if (!string.IsNullOrEmpty(subscriptionId))
            {
                string serviceRootUrl = "https://graph.microsoft.com/beta/subscriptions/";

                // Get an access token and add it to the client.
                string accessToken;
                try
                {
                    string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                    string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
                    SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);
                    string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");

                    AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
                    accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
                }
                catch (Exception ex)
                {
                    return RedirectToAction("Index", "Error", new { message = ex.Message, debug = "" });
                }

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // Send the 'DELETE /subscriptions/id' request.
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, serviceRootUrl + subscriptionId);
                HttpResponseMessage response = await client.SendAsync(request);

                if (!response.IsSuccessStatusCode)
                {
                    return RedirectToAction("Index", "Error", new { message = response.StatusCode, debug = response.Content.ReadAsStringAsync() });
                }
            }
            return RedirectToAction("SignOut", "Account");
        }

    }
}