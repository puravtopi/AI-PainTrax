using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using MS.Models;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class PrintSettingController : Controller
    {
        #region Variables
        private readonly IMapper _mapper;
        private readonly PrintSettingServices _services = new PrintSettingServices();
        private readonly Common _commonservices = new Common();
        private readonly ILogger<PrintSettingController> _logger;
        #endregion

        #region Public Method
        public PrintSettingController(IMapper mapper, ILogger<PrintSettingController> logger)
        {
            _mapper = mapper;
            _logger = logger;
        }

        public IActionResult Index(string searchtxt = "")
        {
            var data = new List<tbl_print_label>();
            try
            {
                string s = "";
                int a = int.Parse(s);
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId).ToString();
                string cnd = "and cmp_id= " + cmpid;
                if (!string.IsNullOrEmpty(searchtxt))
                {
                    cnd = " and title like '%" + searchtxt + "%' ";
                    var result = _services.GetAll();
                    data = result;
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Index");
            }
            return View();
        }

        [HttpPost]
        public IActionResult GetDetails(int Id)
        {
            var data = new tbl_print_label();

            try
            {
                data = _services.GetOne(Id);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "GetDetails");
            }

            return Json(data);
        }

        [HttpPost]
        public IActionResult UpdateDetails(tbl_print_label model)
        {
            try
            {
                _services.Update(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "UpdateDetails");
            }
            return Json(1);
        }


        [HttpPost]
        public IActionResult SaveTemplate(tbl_template model)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid.Value;
                //model.type = "IE";
                _services.SaveTemplate(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SaveTemplate");
            }
            return Json(1);
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
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][name]"].FirstOrDefault();
                // Sort Column Direction ( asc ,desc)
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                // Search Value from (Search box)
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                //Paging Size (10,20,50,100)
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                string cnd = " and lbl_code like '%" + searchValue + "%' or lbl_title like '%" + searchValue + "%' and cmp_id=" + cmpid;
                var Data = _services.GetAll(cnd);

                //Sorting

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


        public IActionResult Templates()
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            ViewBag.visitTypeList = _commonservices.GetVisitTypeForTemplate(cmpid);
            return View();
        }

        public IActionResult TemplateList()
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
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][name]"].FirstOrDefault();
                // Sort Column Direction ( asc ,desc)
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                // Search Value from (Search box)
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                //Paging Size (10,20,50,100)
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                string cnd = " and type like '%" + searchValue + "%' and  cmp_id=" + cmpid;
                var Data = _services.GetAllTemplate(cnd);

                //Sorting

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
                SaveLog(ex, "TemplateList");
                throw;
            }

        }
        public IActionResult ModifyTemplate(int id, string type = "")
        {
            if (type == "")
            {
                var data = _services.GetOneTemplate(id);
                return View(data);
            }
            else
            {
                tbl_template objTemplate = new tbl_template()
                {
                    id=0,
                    type = type
                };
                return View(objTemplate);
            }

        }

        #endregion

        #region private Method
        private void SaveLog(Exception ex, string acctionname)
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
                Message = msg,
            };
            new LogService().Insert(logdata);
        }
        #endregion
    }
}
