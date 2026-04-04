using System.Threading;
using AutoMapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class SettingsController : Controller
    {
        #region Variables
        private readonly IMapper _mapper;
        private readonly SettingsService _services = new SettingsService();
        private readonly ILogger<SettingsController> _logger;
        private readonly Common _commonservices = new Common();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;
        #endregion

        public SettingsController(IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
                               ILogger<LocationController> logger, IConfiguration configuration)
        {
            _mapper = mapper;

            Environment = environment;
            Configuration = configuration;
        }

        public IActionResult Index()
        {
            var data = new tbl_settings();
            try
            {

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
                data = _services.GetOne(cmpid.Value);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Index");
            }

            return View(data);
        }

        [HttpPost]
        public IActionResult Index(tbl_settings model, IFormFile header_template)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid.Value;

                if (header_template != null)
                {
                    string folderPath = Path.Combine(Environment.WebRootPath, "Uploads/HeaderTemplate");
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    // Generate a unique filename 
                    string fileName = Guid.NewGuid().ToString() + Path.GetExtension(header_template.FileName);
                    string filePath = Path.Combine(folderPath, fileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        header_template.CopyTo(fileStream);
                    }
                    model.header_template = fileName;
                }
                else
                {
                    model.header_template = model.header_template_hidden;
                }

                _services.Update(model);

                ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
                HttpContext.Session.SetInt32(SessionKeys.SessionLocationId, model.location.Value);
                HttpContext.Session.SetInt32(SessionKeys.SessionPageSize, model.page_size);
                HttpContext.Session.SetString(SessionKeys.SessionPageBreak, model.pageBreakForInjection.ToString().ToLower());
                HttpContext.Session.SetString(SessionKeys.SessionIsDaignosis, model.isdaignosisshow.ToString().ToLower());
                HttpContext.Session.SetString(SessionKeys.SessionDaignosisFoundStatment, model.foundStatment == null ? "" : model.foundStatment);
                HttpContext.Session.SetString(SessionKeys.SessionDaignosisNotFoundStatment, model.notfoundStatment == null ? "" : model.notfoundStatment);
                HttpContext.Session.SetString(SessionKeys.SessionInjectionAsSeparateBlock, model.injectionAsSeparateBlock.ToString().ToLower());
                HttpContext.Session.SetString(SessionKeys.SessionHeaderTemplate, model.header_template == null ? "" : model.header_template.ToString().ToLower());
                HttpContext.Session.SetString(SessionKeys.SessionFontFamily, model.font_family.ToString().ToLower());
                HttpContext.Session.SetString(SessionKeys.SessionFontSize, model.font_size.ToString().ToLower());
                HttpContext.Session.SetString(SessionKeys.SessionPreop, model.show_preop.ToString());
                HttpContext.Session.SetString(SessionKeys.SessionPostop, model.show_postop.ToString());
                HttpContext.Session.SetString(SessionKeys.SessionGAIT, model.gait_default == null ? "" : model.gait_default.ToString());
                HttpContext.Session.SetString(SessionKeys.SessionFUDate, model.fu_default == null ? "" : model.fu_default.ToString());
                HttpContext.Session.SetString(SessionKeys.SessionSideCase, model.casetype == null ? "U" : model.casetype.ToString());
                HttpContext.Session.SetString(SessionKeys.SessionShowTableBorder, model.table_border.ToString());

                ViewBag.Message = "Settings Updated Successfully";
                return View(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Index");
            }
            return View();
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
