using Microsoft.AspNetCore.Mvc;

namespace PainTrax.Web.Controllers
{
    public class ParentController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
