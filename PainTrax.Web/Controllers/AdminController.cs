using Microsoft.AspNetCore.Mvc;

namespace PainTrax.Web.Controllers
{
    public class AdminController : Controller
    {
        public IActionResult Index()
        {
            ViewBag.ShowSidebar = HttpContext.Session.GetString("IsAdmin") == "true";
            return View();
        }
    }
}
