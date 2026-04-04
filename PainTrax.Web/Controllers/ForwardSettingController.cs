using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using PainTrax.Web.Models;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class ForwardSettingController : Controller
    {

        private readonly IEToFUService _services = new IEToFUService();
        private readonly Common _commonService = new Common();
        private readonly ILogger<ForwardSettingController> _logger;

        public ForwardSettingController(ILogger<ForwardSettingController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult IETOFU()
        
        {
            try
            {
                int cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).Value;
                var data = new IEToFuVMList();

                var soapData = _services.GetAll(" and tbl_name='page1' and cmp_id=" + cmpid);

                var colNames = "";

                foreach (var item in soapData)
                {
                    colNames = colNames + "," + item.tbl_column;
                }

                ViewBag.Page1List = _commonService.GetSOAPSettingList(colNames.TrimStart(',').ToString());

                data.cntSOAP = ViewBag.Page1List.Count;


                var page2Data = _services.GetAll(" and tbl_name='page2' and cmp_id=" + cmpid);

                colNames = "";

                foreach (var item in page2Data)
                {
                    colNames = colNames + "," + item.tbl_column;
                }

                ViewBag.Page2List = _commonService.GetPage2SettingList(colNames.TrimStart(',').ToString());

                data.cntPage2 = ViewBag.Page2List.Count;

                var pageNEData = _services.GetAll(" and tbl_name='ne' and cmp_id=" + cmpid);

                colNames = "";

                foreach (var item in pageNEData)
                {
                    colNames = colNames + "," + item.tbl_column;
                }

                ViewBag.PageNEList = _commonService.GetPageNESettingList(colNames.TrimStart(',').ToString());

                data.cntPageNE = ViewBag.PageNEList.Count;

                var pageOtherData = _services.GetAll(" and tbl_name='other' and cmp_id=" + cmpid);

                colNames = "";

                foreach (var item in pageOtherData)
                {
                    colNames = colNames + "," + item.tbl_column;
                }

                ViewBag.PageOtherList = _commonService.GetPageOtherSettingList(colNames.TrimStart(',').ToString());

                data.cntPageOther = ViewBag.PageOtherList.Count;

                var pagedg = _services.GetAll(" and tbl_name='dg' and cmp_id=" + cmpid);

                colNames = "";

                foreach (var item in pagedg)
                {
                    colNames = colNames + "," + item.tbl_column;
                }

                ViewBag.PageDaignoList = _commonService.GetDaignosisFeilds(colNames.TrimStart(',').ToString());

                data.cntPageDaignoList = ViewBag.PageDaignoList.Count;
                return View(data);
            }
            catch(Exception ex)
            {
                SaveLog(ex, "IETOFU");
                return View("");
            }
        }

        [HttpPost]
        public IActionResult IETOFU(IEToFuVMList model)
        {
            int cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).Value;
            var userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId).Value;

            _services.Delete(cmpid);

            if (model.SOAP != null)
            {
                foreach (var item in model.SOAP)
                {
                    item.created_by = userid;
                    item.cmp_id = cmpid;

                    _services.Insert(item);
                }
            }
            if (model.page2 != null)
            {
                foreach (var item in model.page2)
                {
                    item.created_by = userid;
                    item.cmp_id = cmpid;

                    _services.Insert(item);
                }
            }

            if (model.ne != null)
            {
                foreach (var item in model.ne)
                {
                    item.created_by = userid;
                    item.cmp_id = cmpid;

                    _services.Insert(item);
                }
            }

            if (model.dg != null)
            {
                foreach (var item in model.dg)
                {
                    item.created_by = userid;
                    item.cmp_id = cmpid;

                    _services.Insert(item);
                }
            }


            if (model.other != null)
            {
                foreach (var item in model.other)
                {
                    item.created_by = userid;
                    item.cmp_id = cmpid;

                    _services.Insert(item);
                }

                var other = new IEToFuVM() { 
                    cmp_id = cmpid,
                    created_by=userid,
                    tbl_column= "treatment_delimit",
                    tbl_name= "other"
                };

                _services.Insert(other);
            }
            return Json(1);
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
