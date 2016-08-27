using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Mvc;
using GraphWebhooks.TokenStorage;
using System.Configuration;
using System.Threading.Tasks;
using System.Security.Claims;
using GraphWebhooks.Auth;

namespace GraphWebhooks.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [Authorize]
        [Authorize]
        public async Task<ActionResult> Graph()
        {
            string userObjId = AuthHelper.GetUserId(ClaimsPrincipal.Current);

            RuntimeTokenCache tokenCache = new RuntimeTokenCache(userObjId);

            AuthHelper authHelper = new AuthHelper(tokenCache);

            ViewBag.AccessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
            if (null == ViewBag.AccessToken)
            {
                return new EmptyResult();
            }

            return View();
        }

        [HttpPost]
        public async Task<ActionResult> SendGraphRequest(string accessToken, string requestUrl)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                // Set up the HTTP GET request
                HttpRequestMessage apiRequest = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                apiRequest.Headers.UserAgent.Add(new ProductInfoHeaderValue("OAuthStarter", "1.0"));
                apiRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                apiRequest.Headers.Add("client-request-id", Guid.NewGuid().ToString());
                apiRequest.Headers.Add("return-client-request-id", "true");

                // Send the request and return the JSON body of the response
                HttpResponseMessage response = await httpClient.SendAsync(apiRequest);
                return Json(response.Content.ReadAsStringAsync().Result);
            }
        }
    }
}