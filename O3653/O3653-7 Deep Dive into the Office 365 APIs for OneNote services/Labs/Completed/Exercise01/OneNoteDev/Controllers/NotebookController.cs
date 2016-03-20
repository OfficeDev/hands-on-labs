using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Threading.Tasks;
using OneNoteDev.Auth;
using OneNoteDev.Models;
using OneNoteDev.TokenStorage;

namespace OneNoteDev.Controllers
{
    public class NotebookController : Controller
    {
        [Authorize]
        public async Task<ActionResult> Index()
        {
            // Get an access token for the request.
            string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);
            string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");

            AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
            string accessToken = await authHelper.GetUserAccessToken("/Notebook/Index");

            // Make the request.
            var repository = new NotebookRepository(accessToken);
            var myNotebooks = await repository.GetNotebooks();

            return View(myNotebooks);
        }
    }
}