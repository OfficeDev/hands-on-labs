using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNet.Mvc;
using Microsoft.AspNet.Authorization;

// For more information on enabling MVC for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace GraphFilesWeb.Controllers
{
    public class FilesController : Controller
    {
        [Authorize]
        public async Task<ActionResult> Index(int? pageSize)
        {
            FileRepository repository = new FileRepository(accessToken: null);

            // setup paging defaults if not provided
            pageSize = pageSize ?? 10;

            // setup paging for the IU
            ViewBag.PageSize = pageSize.Value;

            var results = await repository.GetMyFilesAsync(pageSize.Value);
            return View(results);
        }

        [Authorize]
        public async Task<ActionResult> Delete(string id, string etag)
        {
            FileRepository repository = new FileRepository(accessToken: null);
            if (id != null)
            {
                await repository.DeleteItemAsync(id, etag);
            }

            return Redirect("/Files");
        }
    }
}
