using AutoMapper;
using DocumentFormat.OpenXml.Drawing;
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
    public class POCConfigController : Controller
    {
        private readonly ILogger<POCConfigController> _logger;   
        private readonly IMapper _mapper;
        private readonly POCConfigService _services = new POCConfigService();
        private readonly Common _commonservices = new Common();
        private readonly IWebHostEnvironment _environment;

        public POCConfigController(IMapper mapper,ILogger<POCConfigController> logger, IWebHostEnvironment environment)
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
            tbl_pocconfig obj = new tbl_pocconfig();
            var data = new tbl_pocconfig();
            try
            {               
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
                
                obj.Listcolumns = _commonservices.GetpocconfigCheckBoxList(cmpid.Value);


                var downloadFolder = System.IO.Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
                var subFolders = Directory.GetDirectories(downloadFolder);

                var filesByFolder = new Dictionary<string, List<string>>();

                foreach (var folder in subFolders)
                {
                    var folderName = System.IO.Path.GetFileName(folder);
                    var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                            .Select(System.IO.Path.GetFileName)
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
        public IActionResult Create(tbl_pocconfig model)
        {
            try
            {
                tbl_pocconfig obj = new tbl_pocconfig();
                //model.CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                //model.CreatedDate = System.DateTime.Now;
                //model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();

                //obj.Listcolumns = _commonservices.GetpocconfigCheckBoxList(cmpid.Value);
                string loc_ids = "", loc_name = "";

                foreach (var i in model.Listcolumns)
                {
                    if (i.IsChecked)
                    {
                        loc_ids = loc_ids + "," + i.Id;
                        loc_name = loc_name + "," + i.Item;
                    }
                }
                loc_ids = loc_ids.TrimStart(',');
                model.id = loc_name.TrimStart(',');
                model.columns = loc_name.TrimStart(',');
               // model.Listcolumns = _commonservices.GetpocconfigCheckBoxList(cmpid.Value);

                //_services.Insert(model);

                _services.Insert(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return Json(1);
        }
        public IActionResult Edit(int id1)
        {
            var data = new tbl_pocconfig();
            try
            {
                tbl_pocconfig obj = new tbl_pocconfig();
                //obj.id = Convert.toid1;
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
                //data = _services.GetOne(obj);
                data.Listcolumns = _commonservices.GetpocconfigCheckBoxList(cmpid.Value);

               

                var downloadFolder = System.IO.Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
                var subFolders = Directory.GetDirectories(downloadFolder);

                var filesByFolder = new Dictionary<string, List<string>>();

                foreach (var folder in subFolders)
                {
                    var folderName = System.IO.Path.GetFileName(folder);
                    var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                            .Select(System.IO.Path.GetFileName)
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
        public IActionResult Edit(tbl_pocconfig model)
        {
            try
            {
                string loc_ids = "", loc_name = "";
                foreach (var i in model.Listcolumns)
                {
                    if (i.IsChecked)
                    {
                        loc_ids = loc_ids + "," + i.Id;
                        loc_name = loc_name + "," + i.Item.Trim();
                    }
                }
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                loc_ids = loc_ids.TrimStart(',');
                model.id = loc_name.TrimStart(','); ;
                model.columns = loc_name.TrimStart(',');
                model.cmp_id = cmpid;

                _services.Update(model);
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return Json(1);
        }

        public IActionResult ExportEdit(int id1)
        {
            var data = new tbl_pocconfig();
            try
            {
                tbl_pocconfig obj = new tbl_pocconfig();
                //obj.id = Convert.toid1;
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
                //data = _services.GetOne(obj);
                data.Listcolumns = _commonservices.GetpocconfigExportCheckBoxList(cmpid.Value);



                var downloadFolder = System.IO.Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
                var subFolders = Directory.GetDirectories(downloadFolder);

                var filesByFolder = new Dictionary<string, List<string>>();

                foreach (var folder in subFolders)
                {
                    var folderName = System.IO.Path.GetFileName(folder);
                    var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                            .Select(System.IO.Path.GetFileName)
                                            .ToList();

                    filesByFolder.Add(folderName, pdfFiles);
                }

                ViewBag.FilesByFolder = filesByFolder;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult ExportEdit(tbl_pocconfig model)
        {
            try
            {
                string loc_ids = "", loc_name = "";
                foreach (var i in model.Listcolumns)
                {
                    if (i.IsChecked)
                    {
                        loc_ids = loc_ids + "," + i.Id;
                        loc_name = loc_name + "," + i.Item.Trim();
                    }
                }
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                loc_ids = loc_ids.TrimStart(',');
                model.id = loc_name.TrimStart(','); ;
                model.columns = loc_name.TrimStart(',');
                model.export_Columns = loc_name.TrimStart(',');
                model.cmp_id = cmpid;


                _services.ExportUpdate(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return Json(1);
        }


        public IActionResult List()
        {
            try
            {
                var data = new tbl_pocconfig();
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
                string cnd = "";//=  " and cmp_id=" + cmpid + " and (title like '%" + searchValue + "%'  OR Location_ids like '%" + searchValue + "%') ";
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_pocconfig).GetProperties()[Convert.ToInt32(sortColumn)];
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
                var data1 = Data.Skip(skip).Take(pageSize).ToList();
                //Returning Json Data
                return Json(new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data1 });
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
