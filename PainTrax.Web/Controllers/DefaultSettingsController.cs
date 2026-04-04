using Microsoft.AspNetCore.Mvc;
using MS.Models;
using MS.Services;
using Newtonsoft.Json;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class DefaultSettingsController : Controller
    {
        private readonly ILogger<tbl_default> _logger;
        DefaultValueSettingServices services = new DefaultValueSettingServices();
        WebsiteMacrosMasterService _websiteMacrosService = new WebsiteMacrosMasterService();

        public DefaultSettingsController(ILogger<tbl_default> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            SettingsVM model = new SettingsVM();
            try
            {

                var cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);


                var macroList = _websiteMacrosService.GetAutoComplete(cmpid.Value);
                ViewBag.allmacroList = JsonConvert.SerializeObject(macroList);

                macroList = _websiteMacrosService.GetAutoComplete(cmpid.Value, "CC");
                ViewBag.ccmacroList = JsonConvert.SerializeObject(macroList);

                macroList = _websiteMacrosService.GetAutoComplete(cmpid.Value, "PE");
                ViewBag.pemacroList = JsonConvert.SerializeObject(macroList);

                model.Page1 = services.GetOnePage1(cmpid.Value);

                if (model.Page1 == null)
                    model.Page1 = new tbl_ie_page1_default();

                model.Page2 = services.GetOnePage2(cmpid.Value);

                if (model.Page2 == null)
                    model.Page2 = new tbl_ie_page2_default();

                model.NE = services.GetOneNE(cmpid.Value);
                if (model.NE == null)
                {
                    model.NE = new tbl_ie_ne_default();
                    model.NE.id = 0;
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Index");
            }
            return View(model);
        }

        [HttpPost]
        public IActionResult SavePage1(tbl_ie_page1_default model)
        {
            var data = 0;
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;
                if (model.id > 0)
                {
                    data = model.id.Value;
                    services.UpdatePage1(model);
                }
                else
                    data = services.InsertPage1(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePage1");
            }

            return Json(data);
        }

        [HttpPost]
        public IActionResult SavePage2(tbl_ie_page2_default model)
        {
            var data = 0;
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;

                if (model.id > 0)
                {
                    data = model.id.Value;
                    services.UpdatePage2(model);
                }
                else
                    data = services.InsertPage2(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePage2");
            }
            return Json(data);
        }

        [HttpPost]
        public IActionResult SaveNE(tbl_ie_ne_default model)
        {
            var data = 0;
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;

                if (model.id > 0)
                {
                    data = model.id;
                    services.UpdateNE(model);
                }
                else
                    data = services.InsertNE(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SaveNE");
            }
            return Json(data);
        }

        #region public Method
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
                Message = msg,
            };
        }
        #endregion
    }
}
