using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class CompanyController : Controller
    {
        private readonly IMapper _mapper;
        private readonly ILogger<tbl_company> _logger;
        private readonly CompanyServices _services = new CompanyServices();

        public CompanyController(IMapper mapper,ILogger<tbl_company> logger)
        {
            _mapper = mapper;
            _logger = logger;
        }
        public IActionResult Index(string searchtxt = "")
        {          
            return View();
        }

        public IActionResult Create()
        {
            tbl_company obj = new tbl_company();
            try
            {                       
                obj.is_active = false;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }            
            return View(obj);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(tbl_company model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    model.password = EncryptionHelper.Encrypt(model.password);
                    _services.Insert(model);
                }
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Create");
            }            
            return RedirectToAction("Index");
        }

        public IActionResult Edit(int id)
        {
            tbl_company data = new tbl_company();
            try
            {
                tbl_company obj = new tbl_company();
                obj.id = id;
                data = _services.GetOne(obj);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_company model)
        {
            try
            {
                _services.Update(model);
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Delete(int id)
        {
            try
            {
                tbl_company obj = new tbl_company();
                obj.id = id;
                _services.Delete(obj);
            }
            catch(Exception ex)
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
                string cnd = "";
                if (!String.IsNullOrEmpty(searchValue))
                {
                    cnd = " and (name like '%" + searchValue + "%' or address like '%" + searchValue + "%' or email like '%" + searchValue + "%' or telephone like '%" + searchValue + "%'  ";
                    
                }
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_company).GetProperties()[Convert.ToInt32(sortColumn)];
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

        #region Private Method
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
