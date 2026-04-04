using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MS.Models;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class VisitTypeController : Controller
    {
        private readonly IMapper _mapper;
        private readonly ILogger<VisitTypeController> _logger;
        private readonly Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private readonly VisitTypeService _services = new VisitTypeService();

        public VisitTypeController(ILogger<VisitTypeController> logger, IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _logger = logger;
            _mapper = mapper;
            Environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Create()
        {
            tbl_visit_type obj = new tbl_visit_type();
            return View(obj);
        }
        [HttpPost]
        [AllowAnonymous]
        public IActionResult Create(tbl_visit_type model)
        {
            try
            {
                if (ModelState.IsValid)
                {

                    string cnd = " and type like '%" + model.type + "%' ";
                    var Data = _services.GetAll(cnd);

                    if (Data.Count > 0) {
                        ViewBag.error = "This visit type is already present.";
                        return View(model);
                    }
                    _services.Insert(model);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return RedirectToAction("Index");
        }
        public IActionResult Edit(int id)
        {
            tbl_visit_type data = new tbl_visit_type();
            try
            {
                tbl_visit_type obj = new tbl_visit_type();
                obj.id = id;
                data = _services.GetOne(id);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }
        [HttpPost]
        public IActionResult Edit(tbl_visit_type model)
        {
            _services.Update(model);
            return RedirectToAction("Index");
        }
        public IActionResult Delete(int id)
        {
            try
            {
                tbl_visit_type obj = new tbl_visit_type();
                obj.id = id;
                //_services.Delete(obj);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Delete");
            }
            return RedirectToAction("Index");
        }
        public IActionResult List()
        {
            try
            {

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
                string cnd = " and type like '%" + searchValue + "%' ";
                var Data = _services.GetAll(cnd);

                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_visit_type).GetProperties()[Convert.ToInt32(sortColumn)];
                    if (sortColumnDirection.ToUpper() == "ASC")
                    {
                        Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                    }
                    else
                    {
                        Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                    }
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
                SaveLog(ex, "List");
            }
            return Json("");
        }
        #region Private Method
        private void SaveLog(Exception ex, string actionname)
        {
            var msg = "";
            if (ex.InnerException == null)
            {
                _logger.LogError(ex.Message);
                msg = ex.Message;
            }
            else
            {
                _logger.LogError(ex.InnerException.Message);
                msg = ex.InnerException.Message;
            }
            var logdata = new tbl_log
            {
                CreatedDate = DateTime.Now,
                CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId),
                Message = msg
            };
            new LogService().Insert(logdata);
        }
        #endregion
    }
}
