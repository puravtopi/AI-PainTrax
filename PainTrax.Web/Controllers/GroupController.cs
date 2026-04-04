using AutoMapper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Globalization;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class GroupController : Controller
    {
        private readonly ILogger<GroupController> _logger;   
        private readonly IMapper _mapper;
        private readonly GroupsService _services = new GroupsService();
        private readonly Common _commonservices = new Common();
        private readonly IWebHostEnvironment _environment;

        public GroupController(IMapper mapper,ILogger<GroupController> logger, IWebHostEnvironment environment)
        {
            _mapper = mapper;
            _logger = logger;
            _environment = environment;
        }

        public IActionResult Index()
        {       
            return View();
        }

        public IActionResult Create()
        {
            tbl_groups obj = new tbl_groups();
            try
            {               
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
                obj.LocationList = _commonservices.GetLocationsCheckBoxList(cmpid.Value);
                obj.PagesList = _commonservices.GetPagesCheckBoxList();
                obj.ReportsList = _commonservices.GetReportsCheckBoxList();
                obj.RoleList = _commonservices.GetRolsCheckBoxList();

                var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
                var subFolders = Directory.GetDirectories(downloadFolder);

                var filesByFolder = new Dictionary<string, List<string>>();

                foreach (var folder in subFolders)
                {
                    var folderName = Path.GetFileName(folder);
                    var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                            .Select(Path.GetFileName)
                                            .ToList();

                    filesByFolder.Add(folderName, pdfFiles);
                }

                ViewBag.FilesByFolder = filesByFolder;
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Create");
            }
             return View(obj);
        }

        [HttpPost]
        public IActionResult Create(tbl_groups model)
        {
            try
            {                
                model.CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.CreatedDate = System.DateTime.Now;
                model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                string loc_ids = "", loc_name = "";

                foreach (var i in model.LocationList)
                {
                    if (i.IsChecked)
                    {
                        loc_ids = loc_ids + "," + i.Id;
                        loc_name = loc_name + "," + i.Item;
                    }
                }
                loc_ids = loc_ids.TrimStart(',');
                model.Location_ids = loc_ids;
                model.location_name = loc_name.TrimStart(',');

                string page_ids = "", page_name = "";

                foreach (var i in model.PagesList)
                {
                    if (i.IsChecked)
                    {
                        page_ids = page_ids + "," + i.Id;
                        page_name = page_name + "," + i.Item;
                    }
                }
                model.pages_ids = page_ids.TrimStart(',');
                model.pages_name = page_name.TrimStart(',');
                string report_ids = "", report_name = "";
                
                foreach (var i in model.ReportsList)
                {
                    if (i.IsChecked)
                    {
                        report_ids = report_ids + "," + i.Id;
                        report_name = report_name + "," + i.Item;
                    }
                }
                model.reports_ids = report_ids.TrimStart(',');
                model.report_name = report_name.TrimStart(',');
                string role_ids = "", role_name = "";

                foreach (var i in model.RoleList)
                {
                    if (i.IsChecked)
                    {
                        role_ids = role_ids + "," + i.Id;
                        role_name = role_name + "," + i.Item;
                    }
                }
                model.role_ids = role_ids.TrimStart(',');
                model.role_name = role_name.TrimStart(',');


                string form_name = "";

                foreach (var i in model.FormList)
                {
                    if (i.IsChecked)
                    {
                      
                        form_name = form_name + "," + i.Item;
                    }
                }
             
                model.form_name = form_name.TrimStart(',');
                _services.Insert(model);

                //_services.Insert(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return Json(1);
        }
        public IActionResult Edit(int id)
        {
            var data = new tbl_groups();
            try
            {
                tbl_groups obj = new tbl_groups();
                obj.Id = id;
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
                data = _services.GetOne(obj);
                data.LocationList = _commonservices.GetLocationsCheckBoxList(cmpid.Value, data.Location_ids);
                data.PagesList = _commonservices.GetPagesCheckBoxList(data.pages_ids);
                data.ReportsList = _commonservices.GetReportsCheckBoxList(data.reports_ids);
                data.RoleList = _commonservices.GetRolsCheckBoxList(data.role_ids);
                data.form_name = data.form_name == null ? "" : data.form_name;

                var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
                var subFolders = Directory.GetDirectories(downloadFolder);

                var filesByFolder = new Dictionary<string, List<string>>();

                foreach (var folder in subFolders)
                {
                    var folderName = Path.GetFileName(folder);
                    var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                            .Select(Path.GetFileName)
                                            .ToList();

                    filesByFolder.Add(folderName, pdfFiles);
                }

                ViewBag.FilesByFolder = filesByFolder;
            } 
            catch(Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_groups model)
        {
            try
            {
                string loc_ids = "", loc_name = "";
                foreach (var i in model.LocationList)
                {
                    if (i.IsChecked)
                    {
                        loc_ids = loc_ids + "," + i.Id;
                        loc_name = loc_name + "," + i.Item.Trim();
                    }
                }
                loc_ids = loc_ids.TrimStart(',');
                model.Location_ids = loc_ids;
                model.location_name = loc_name.TrimStart(',');
                string page_ids = "", page_name = "";
                foreach (var i in model.PagesList)
                {
                    if (i.IsChecked)
                    {
                        page_ids = page_ids + "," + i.Id;
                        page_name = page_name + "," + i.Item;
                    }
                }
                model.pages_ids = page_ids.TrimStart(',');
                model.pages_name = page_name.TrimStart(',');

                string report_ids = "", report_name = "";

                foreach (var i in model.ReportsList)
                {
                    if (i.IsChecked)
                    {
                        report_ids = report_ids + "," + i.Id;
                        report_name = report_name + "," + i.Item;
                    }
                }
                model.reports_ids = report_ids.TrimStart(',');
                model.report_name = report_name.TrimStart(',');

                string role_ids = "", role_name = "";
                foreach (var i in model.RoleList)
                {
                    if (i.IsChecked)
                    {
                        role_ids = role_ids + "," + i.Id;
                        role_name = role_name + "," + i.Item;
                    }
                }
                model.role_ids = role_ids.TrimStart(',');
                model.role_name = role_name.TrimStart(',');

                string form_name = "";

                if (model.FormList != null)
                {
                    foreach (var i in model.FormList)
                    {
                        if (i.IsChecked)
                        {

                            form_name = form_name + "," + i.Item;
                        }
                    }
                }

                model.form_name = form_name.TrimStart(',');

                _services.Update(model);
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return Json(1);
        }
        public IActionResult Delete(int id)
        {
            tbl_groups obj = new tbl_groups();
            try
            {
                obj.Id = id;
                _services.Delete(obj);
            }
            catch(Exception ex)
            {
                SaveLog(ex,"Delete ");
            }
            return RedirectToAction("Index");
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
                string cnd =  " and cmp_id=" + cmpid + " and (title like '%" + searchValue + "%'  OR Location_ids like '%" + searchValue + "%') ";
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_groups).GetProperties()[Convert.ToInt32(sortColumn)];
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
                SaveLog(ex, "List");
            }
            return Json("");
        }

        #region private Method
        private void SaveLog(Exception ex, string actionname)
        {
            var msg = "";
            if(ex.InnerException == null)
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
                Message = msg,
            };
            new LogService().Insert(logdata);
        }
        #endregion 
    }
}
