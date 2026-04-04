using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class FollowupFormController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Create()
        {
            FollowupForm obj = new FollowupForm();
           // obj.dov = Convert.ToDateTime(System.DateTime.Now.ToString("MM/dd/yyyy"));
            obj.dov = Convert.ToDateTime(System.DateTime.Now.ToShortDateString());
            return View(obj);
        }
        [HttpPost]
        public IActionResult Create(FollowupForm model)
        {




            return RedirectToAction("Create", "FollowupForm");
        }
    }
}
