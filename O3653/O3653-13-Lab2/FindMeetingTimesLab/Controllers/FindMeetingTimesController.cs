using System;
using System.Collections.Generic;
using System.Web.Mvc;
using FindMeetingTimesLab.TokenStorage;
using System.Configuration;
using System.Threading.Tasks;
using FindMeetingTimesLab.Auth;

namespace FindMeetingTimesLab.Controllers
{
    public class FindMeetingTimesController : Controller
    {
        // GET: FindMeetingTimes
        public async Task<ActionResult> Index(string attendees)
        {
            string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

            string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");

            AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
            string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));

            if (!string.IsNullOrEmpty((string)TempData["error"]))
            {
                ViewBag.ErrorMessage = (string)TempData["error"];
            }

            //For first time load, just load the form and the table with no results 
            if (this.Request.HttpMethod == "GET")
            {
                return View();
            }
            
            try
            {
                var client = new GraphHelper();
                client.anchorMailbox = (string)Session["user_name"];
                ViewBag.UserName = client.anchorMailbox;
                string payload = client.GeneratePayload(attendees);

                var results = await client.GetMeetingTimes(accessToken, client.anchorMailbox, payload);

                return View(results);
            }
            catch (Exception ex)
            {
                return RedirectToAction("Index", "Error", new { message = ex.Message });
            }
        }
    }
}