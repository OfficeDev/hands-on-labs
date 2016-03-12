using System.Web.Mvc;

namespace InboxSync.Controllers
{
  public class ErrorController : Controller
  {
    // GET: Error
    public ActionResult Index(string message, string debug)
    {
      ViewBag.Message = message;
      ViewBag.Debug = debug;
      return View("Error");
    }
  }
}