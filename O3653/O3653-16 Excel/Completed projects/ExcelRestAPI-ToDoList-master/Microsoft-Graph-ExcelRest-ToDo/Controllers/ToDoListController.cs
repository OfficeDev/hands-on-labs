using System.Collections.Generic;
using System.Web.Mvc;
using System.Threading.Tasks;
using System;
using Microsoft_Graph_ExcelRest_ToDo.TokenStorage;
using Microsoft_Graph_ExcelRest_ToDo.Auth;
using System.Configuration;

namespace Microsoft_Graph_ExcelRest_ToDo.Controllers
{
    public class ToDoListController : Controller
    {

        //
        // GET: ToDoList
        public async Task<ActionResult> Index()
        {
            string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

            string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");
            AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
            string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));

            await RESTAPIHelper.LoadWorkbook(accessToken);

            return View(await RESTAPIHelper.GetToDoItems(accessToken));
        }

        // GET: ToDoList/Create
        public ActionResult Create()
        {
            var priorityList = new SelectList(new[]
                                          {
                                              new {ID="1",Name="High"},
                                              new{ID="2",Name="Normal"},
                                              new{ID="3",Name="Low"},
                                          },
                            "ID", "Name", 1);
            ViewData["priorityList"] = priorityList;

            var statusList = new SelectList(new[]
                              {
                                              new {ID="1",Name="Not started"},
                                              new{ID="2",Name="In-progress"},
                                              new{ID="3",Name="Completed"},
                                          },
                "ID", "Name", 1);
            ViewData["statusList"] = statusList;

            return View();
        }

        // POST: ToDoList/Create
        [HttpPost]
        public async Task<ActionResult> Create(FormCollection collection)
        {
            try
            {

                string userObjId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

                string tenantId = System.Security.Claims.ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
                string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], tenantId, "");
                AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"], ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);
                string accessToken = await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));

                await RESTAPIHelper.CreateToDoItem(
                    accessToken,
                    collection["Title"],
                    collection["PriorityDD"],
                    collection["StatusDD"],
                    collection["PercentComplete"],
                    collection["StartDate"],
                    collection["EndDate"],
                    collection["Notes"]);
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

    }
}
