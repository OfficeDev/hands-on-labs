using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Graph;
using GraphFilesWeb.Auth;
using GraphFilesWeb.TokenStorage;
using System.Configuration;
using System.Threading.Tasks;

namespace GraphFilesWeb.Controllers
{
    public class FilesController : Controller
    {

        private GraphServiceClient GetGraphServiceClient()
        {
            string userObjId = AuthHelper.GetUserId(System.Security.Claims.ClaimsPrincipal.Current);
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common", "");

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
            }));

            return client;
        }
        
        [Authorize]
        public async Task<ActionResult> Index(int? pageSize, string nextLink)
        {
            var client = GetGraphServiceClient();

            pageSize = pageSize ?? 25;

            IChildrenCollectionRequest request = client.Me.Drive.Root.Children.Request().Top(pageSize.Value);
            if (nextLink != null)
            {
                request = new ChildrenCollectionRequest(nextLink, client, null);
            }

            var results = await request.GetAsync();
            if (null != results.NextPageRequest)
            {
                ViewBag.NextLink = results.NextPageRequest.GetHttpRequestMessage().RequestUri;
            }
            else
            {
                ViewBag.NextLink = null;
            }

            return View(results);
        }

        [Authorize]
        public async Task<ActionResult> Delete(string itemId, string etag)
        {
            var client = GetGraphServiceClient();

            // Build a request and set the If-Match header with the etag
            var request = client.Me.Drive.Items[itemId].Request(new List<Option> { new HeaderOption("If-Match", etag) });

            // Submit the delete request
            await request.DeleteAsync();

            return Redirect("/Files");
        }

        [Authorize]
        public async Task<ActionResult> Upload()
        {
            var client = GetGraphServiceClient();

            foreach (string key in Request.Files)
            {
                var fileInRequest = Request.Files[key];
                if (fileInRequest != null && fileInRequest.ContentLength > 0)
                {
                    var filename = System.IO.Path.GetFileName(fileInRequest.FileName);
                    var request = client.Me.Drive.Root.Children[filename].Content.Request();
                    var createdFile = await request.PutAsync<DriveItem>(fileInRequest.InputStream);
                }
            }

            return Redirect("/Files");
        }


    }
}