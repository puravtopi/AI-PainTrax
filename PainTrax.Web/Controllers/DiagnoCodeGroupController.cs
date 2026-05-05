using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    public class DiagnoCodeGroupController : Controller
    {
        private readonly DiagcodesService _services = new DiagcodesService();
        private readonly Common _common = new Common();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;
        private readonly ILogger<DiagcodeController> _logger;

        public DiagnoCodeGroupController(Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
            IConfiguration configuration, ILogger<DiagcodeController> logger)
        {
            _logger = logger;
            Environment = environment;
            Configuration = configuration;
        }


        public IActionResult Index()
        {
            return View();
        }
        public IActionResult List()
        {
            try
            {
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                var draw = HttpContext.Request.Form["draw"].FirstOrDefault();
                // Skiping number of Rows count
                var start = Request.Form["start"].FirstOrDefault();
                // Paging Length 10,20
                var length = Request.Form["length"].FirstOrDefault();
                // Sort Column Name
                var sortColumn = Request.Form["order[0][column]"].FirstOrDefault();
                // Sort Column Direction ( asc ,desc)
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                // Search Value from (Search box)
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                //Paging Size (10,20,50,100)
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                string cnd = "  and cmp_id=" + cmpid + "  and (bodypart like '%" + searchValue + "%' or groupname like '%" + searchValue + "%')";
                var Data = _services.GetAllDiagCodeGroups(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_diagcodes_group).GetProperties()[Convert.ToInt32(sortColumn)];
                    if (sortColumnDirection.ToUpper() == "ASC")
                    {
                        Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                    }
                    else
                        Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                }
                //Search
                //total number of rows count 
                recordsTotal = Data.Count();
                //Paging 
                var data = Data.Skip(skip).Take(pageSize).ToList();
                //Returning Json Data
                return Json(new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data });

            }
            catch (Exception ex)
            {
                //SaveLog(ex, "List");
            }
            return Json("");
        }
        public IActionResult Create()
        {
            tbl_diagcodes_group obj = new tbl_diagcodes_group();
            ViewBag.isError = false;
            return View(obj);
        }

        [HttpPost]
        [AllowAnonymous]
        public IActionResult Create(tbl_diagcodes_group model)
        {
            try
            {
                if (ModelState.IsValid)
                {

                    model.Cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    //model.PreSelect =null;
                    _services.InsertDiagCodeGroup(model);
                }
            }
            catch (Exception ex)
            {

            }
            return RedirectToAction("Index");
        }

        public IActionResult Edit(int id)
        {
            tbl_diagcodes_group data = new tbl_diagcodes_group();
            try
            {
                ViewBag.isError = false;
                var obj = _services.GetAllDiagCodeGroups(" and id=" + id).FirstOrDefault();
                data =obj != null ? obj : new tbl_diagcodes_group();
            }
            catch (Exception ex)
            {

            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_diagcodes_group model)
        {
            try
            {
                _services.UpdateDiagCodeGroup(model);
            }
            catch (Exception ex)
            {

            }
            return RedirectToAction("Index");
        }
    }
}
