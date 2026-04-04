using AutoMapper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Diagnostics;
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;

namespace PainTrax.Web.Controllers
{

    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IMapper _mapper;
        private readonly LoginServices _services = new LoginServices();
        private readonly GroupsService _groupservices = new GroupsService();
        private readonly DashboardService _dashboardservices = new DashboardService();
        private readonly ISession _session = null;
        private readonly IHttpContextAccessor _httpContextAccessor = null;
        private readonly Common _commonservices = new Common();
        private readonly UserService _userService = new UserService();
        private readonly SettingsService _setting = new SettingsService();
        private readonly IEmailService _emailService;
        private readonly POCServices _pocservices = new POCServices();
        private readonly SurgeryCentreService _surgeryCentreService = new SurgeryCentreService();
        private readonly DataTransferService _dataTransferService;
      



        public HomeController(ILogger<HomeController> logger,
            IEmailService emailService, DataTransferService dataTransferService,
            IMapper mapper, IHttpContextAccessor httpContextAccessor
          )
        {
            _logger = logger;
            _mapper = mapper;
            _httpContextAccessor = httpContextAccessor;
            _session = _httpContextAccessor.HttpContext.Session;
            _emailService = emailService;
            _dataTransferService = dataTransferService;
            

        }

        public IActionResult Index()
        {

            // var _data = _dataTransferService.UpdateInsurance();
            // this code is use to delete unwanted mcode from procedure table

            //// Static path (adjust as per your environment)
            //string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "poc.xlsx");

            //var list = new List<MyPOCExcelModel>();

            //using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            //{
            //    WorkbookPart workbookPart = document.WorkbookPart;
            //    Sheet firstSheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
            //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id);
            //    Worksheet worksheet = worksheetPart.Worksheet;
            //    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            //    var rows = sheetData.Elements<Row>();

            //    foreach (Row row in rows.Skip(1)) // skip header
            //    {
            //        var cells = row.Elements<Cell>().ToList();


            //        var mcode = GetCellValue(workbookPart, cells[0]);

            //        _pocservices.deleteMCode(mcode);


            //    }
            //}


            //TblCompany tbl = new TblCompany()
            //{
            //    Address = "123",
            //    CreatedBy = 1,
            //    CreatedDate = DateTime.Now,
            //    Email = "purav.topi@gmail.com",
            //    Id = 1,
            //    IsActive = true,
            //    Name = "Purav",
            //    Phone = "789"

            //};

            //var t = _mapper.Map<CompanyVM>(tbl);

            //var respose = _service.Save(t);

            var c_id = EncryptionHelper.Encrypt("10");
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                ViewBag.TotalPatient = _dashboardservices.GetTotalPatient(cmpid.Value, 4);
                ViewBag.TodaysPatient = _dashboardservices.GetTotalPatient(cmpid.Value, 1);
                ViewBag.TodaysAttorny = _dashboardservices.GetTotalAttorny(cmpid.Value);
                ViewBag.TodaysIncCo = _dashboardservices.GetTotalInsuranceCompany(cmpid.Value);
                ViewBag.TotalClaimNo = _dashboardservices.GetTotalClaimNo(cmpid.Value);

                string cnd = " and cmp_id=" + cmpid;
                var data = _surgeryCentreService.GetAll(cnd);
                var list = new List<SelectListItem>();



                foreach (var item in data)
                {
                    list.Add(new SelectListItem
                    {
                        Text = item.Surgerycenter_name.ToString(),
                        Value = item.Id.ToString()
                    });
                }
                ViewBag.surgoryList = list;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Index");
            }
            return View();
        }

        public IActionResult ScheduleList(string fDate = "", string tDate = "")
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



                DateTime fdate = string.IsNullOrEmpty(fDate) ? new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1) : Convert.ToDateTime(fDate);
                DateTime tdate = string.IsNullOrEmpty(tDate) ? fdate.AddMonths(1).AddDays(-1) : Convert.ToDateTime(tDate);


                var Data = _pocservices.GetSurgoryDashboardData(fdate.ToString("yyyy-MM-dd"), tdate.ToString("yyyy-MM-dd"), cmpid.ToString());

                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(SurgoryCenterDashboardVM).GetProperties()[Convert.ToInt32(sortColumn)];
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

        private string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
                           .ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }

        public async Task<IActionResult> Login()
        {

            //string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads/DriverLicense/Capture.png");

            ////var service = new OcrService();
            ////var data = service.ExtractData(path);

            //var text = await _azureService.ExtractTextFromImage(path);
            //var data = _azureService.ParseData(text);


            if (Request.Cookies["LoginCookie"] != null)
            {
                var savedLogin = JsonConvert.DeserializeObject<LoginVM>(Request.Cookies["LoginCookie"]);

                // Pre-populate login fields with saved credentials
                return View(savedLogin);
            }
            ViewBag.Success = true;
            return View();
        }

        [HttpPost]
        public IActionResult Login(LoginVM login, bool remember)
        {
            try
            {
                var response = _services.CompanyLogin(login);


                if (response.Success)
                {

                    HttpContext.Session.SetInt32(SessionKeys.SessionCmpId, response.Model.cmp_id.Value);
                    HttpContext.Session.SetString(SessionKeys.SessionCmpClientId, response.Model.client_code);
                    HttpContext.Session.SetString(SessionKeys.SessionCmpName, response.Model.company_name);
                    HttpContext.Session.SetInt32(SessionKeys.SessionCmpUserId, response.Model.Id.Value);
                    HttpContext.Session.SetString(SessionKeys.SessionCmpEmail, response.Model.emailid);
                    HttpContext.Session.SetString(SessionKeys.SessionUserName, response.Model.uname);
                    HttpContext.Session.SetString(SessionKeys.SessionDesignation, response.Model.desig_name);
                    HttpContext.Session.SetInt32(SessionKeys.SessionUserId, response.Model.Id.Value);

                    var setting = _setting.GetOne(response.Model.cmp_id.Value);

                    if (setting != null)
                    {
                        HttpContext.Session.SetInt32(SessionKeys.SessionLocationId, setting.location.Value);
                        HttpContext.Session.SetInt32(SessionKeys.SessionPageSize, setting.page_size);
                        HttpContext.Session.SetString(SessionKeys.SessionDateFormat, setting.dateformat == null ? "MM/dd/yyyy" : setting.dateformat);
                        HttpContext.Session.SetString(SessionKeys.SessionPageBreak, setting.pageBreakForInjection.ToString().ToLower());
                        HttpContext.Session.SetString(SessionKeys.SessionIsDaignosis, setting.isdaignosisshow.ToString().ToLower());
                        HttpContext.Session.SetString(SessionKeys.SessionDaignosisFoundStatment, setting.foundStatment == null ? "" : setting.foundStatment);
                        HttpContext.Session.SetString(SessionKeys.SessionDaignosisNotFoundStatment, setting.notfoundStatment == null ? "" : setting.notfoundStatment);
                        HttpContext.Session.SetString(SessionKeys.SessionInjectionAsSeparateBlock, setting.injectionAsSeparateBlock.ToString().ToLower());
                        HttpContext.Session.SetString(SessionKeys.SessionHeaderTemplate, string.IsNullOrEmpty(setting.header_template) ? "" : setting.header_template.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionPostop, setting.show_postop == null ? "true" : setting.show_postop.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionPreop, setting.show_preop == null ? "true" : setting.show_preop.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionGAIT, setting.gait_default == null ? "" : setting.gait_default.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionFUDate, setting.fu_default == null ? "" : setting.fu_default.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionSideCase, setting.casetype == null ? "U" : setting.casetype.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionShowTableBorder, setting.table_border.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionFontFamily, setting.font_family == null ? "" : setting.font_family.ToString());
                        HttpContext.Session.SetString(SessionKeys.SessionFontSize, setting.font_size == null ? "" : setting.font_size.ToString());

                    }
                    else
                    {
                        HttpContext.Session.SetInt32(SessionKeys.SessionLocationId, 0);
                        HttpContext.Session.SetInt32(SessionKeys.SessionPageSize, 25);
                        HttpContext.Session.SetString(SessionKeys.SessionDateFormat, "MM/dd/yyyy");
                        HttpContext.Session.SetString(SessionKeys.SessionPageBreak, "false");
                        HttpContext.Session.SetString(SessionKeys.SessionIsDaignosis, "false");
                        HttpContext.Session.SetString(SessionKeys.SessionHeaderTemplate, "");
                        HttpContext.Session.SetString(SessionKeys.SessionPreop, "true");
                        HttpContext.Session.SetString(SessionKeys.SessionPostop, "true");
                        HttpContext.Session.SetString(SessionKeys.SessionSideCase, "U");
                        HttpContext.Session.SetString(SessionKeys.SessionShowTableBorder, "true");

                    }


                    //get menu access from assign group

                    tbl_groups objgroup = new tbl_groups() { Id = response.Model.groupid };
                    var groupDetails = _groupservices.GetOne(objgroup);

                    if (groupDetails != null)
                    {
                        HttpContext.Session.SetString(SessionKeys.SessionPagesAccess, groupDetails.pages_name);
                        HttpContext.Session.SetString(SessionKeys.SessionRoleAccess, groupDetails.role_name);
                        HttpContext.Session.SetString(SessionKeys.SessionFormsAccess, groupDetails.form_name == null ? "" : groupDetails.form_name);
                    }


                    // Check if "Remember Me" checkbox is checked
                    if (remember)
                    {
                        // Set a cookie with user's login information
                        var options = new CookieOptions
                        {
                            Expires = DateTime.Now.AddDays(7) // Cookie expires in 7 days
                        };
                        Response.Cookies.Append("LoginCookie", JsonConvert.SerializeObject(login), options);
                    }
                    if (response.Model != null)
                    {
                        if ((bool)response?.Model?.is_newuser)
                        {
                            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                            string tqrs = EncryptionHelper.Encrypt(response?.Model?.Id.ToString());

                            return RedirectToAction("ChangePassword", new { tqrs = tqrs });
                        }

                    }

                    return RedirectToAction("GetProvider");
                }
                else
                {
                    ViewBag.Success = false;
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Login");
            }
            return View(login);
        }

        public IActionResult AdminLogin()
        {
            ViewBag.Success = true;
            return View();
        }

        [HttpPost]
        public IActionResult AdminLogin(LoginVM login)
        {
            try
            {
                var response = _services.AdminLogin(login);

                if (response.Success)
                {
                    HttpContext.Session.SetInt32(SessionKeys.SessionCmpUserId, response.Model.id);
                    HttpContext.Session.SetString(SessionKeys.SessionUserName, response.Model.uname);
                    return RedirectToAction("Index", "Admin");
                }
                else
                {
                    ViewBag.Success = false;
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "AdminLogin");
            }
            return View(login);
        }

        public IActionResult AdminDashboard()
        {
            ViewBag.Success = true;
            return View();
        }

        public IActionResult GetProvider()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            //ViewBag.locList
            var data = _commonservices.GetLocations(cmpid.Value);

            List<SelectListItem> lst = new List<SelectListItem>();

            int defaultlocation = HttpContext.Session.GetInt32(SessionKeys.SessionLocationId).Value;

            foreach (var item in data)
            {
                var obj = new SelectListItem()
                {
                    Text = item.Text,
                    Value = item.Value,
                    Selected = item.Value == defaultlocation.ToString() ? true : false
                };
                lst.Add(obj);

            }
            ViewBag.locList = lst;

            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;


            if (HttpContext.Session.GetString(SessionKeys.SessionDesignation) == "Provider")
            {
                int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionUserId);
                HttpContext.Session.SetInt32(SessionKeys.SessionSelectedProviderId, userid == null ? 0 : userid.Value);
            }

            return View();
        }

        [HttpPost]
        public IActionResult GetProvider(GetProviderVM model, [FromForm] string[] selectedProviders)
        {


            HttpContext.Session.SetInt32(SessionKeys.SessionLocationId, model.locationid == null ? 0 : model.locationid.Value);
            if (HttpContext.Session.GetString(SessionKeys.SessionDesignation) != "Provider")
            {
                HttpContext.Session.SetInt32(SessionKeys.SessionSelectedProviderId, model.providerid == null ? 0 : model.providerid.Value);
            }

            return RedirectToAction("Index", "Home");
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult Logout()
        {
            _session.Clear();
            /* foreach (var item in _httpContextAccessor.HttpContext.Request.Cookies.Keys)
             {
                 _httpContextAccessor.HttpContext.Response.Cookies.Delete(item);
             }*/
            return RedirectToAction("Login", "Home");
        }
        public IActionResult AdminLogout()
        {
            _session.Clear();
            /* foreach (var item in _httpContextAccessor.HttpContext.Request.Cookies.Keys)
             {
                 _httpContextAccessor.HttpContext.Response.Cookies.Delete(item);
             }*/
            return RedirectToAction("AdminLogin", "Home");
        }

        [HttpPost]
        public IActionResult TotalPatient(int type)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var cnt = _dashboardservices.GetTotalPatient(cmpid.Value, type);
            return Json(cnt);
        }

        public IActionResult SessionExpired()
        {
            return View();
        }

        [HttpGet]
        public IActionResult CheckSession()
        {
            bool isExpired = string.IsNullOrEmpty(HttpContext.Session.GetString(SessionKeys.SessionUserName)); // Change as per your session key
            //bool isExpired = false;
            return Json(new { sessionExpired = isExpired });
        }

        [HttpGet]
        public async Task<IActionResult> ForgotPassword()
        {
            ViewBag.Success = false;
            ViewBag.Error = false;
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ForgotPassword(PainTrax.Web.ViewModel.ForgotPassword model)
        {


            var user = _userService.GetOneByEmail(model.email, model.companycode);

            if (user != null)
            {
                var subject = "Reset Your Password";
                var link = "https://paintrax.com/v2/home/ResetPassword?tqrs=" + EncryptionHelper.Encrypt(user.Id.ToString() + "");
                var body = System.IO.File.ReadAllText("wwwroot/Uploads/EmailTemplate/ForgotPassword.html")
                               .Replace("{RESET_LINK}", link);




                await _emailService.SendEmailAsync("purav.topi@gmail.com", subject, body);
                ViewBag.Success = true;
                ViewBag.Error = false;
            }
            else
            {
                ViewBag.Success = false;
                ViewBag.Error = true;
            }
            return View();
        }

        [HttpGet]
        public IActionResult ResetPassword(string tqrs)
        {
            tqrs = tqrs.Replace(" ", "+");
            var model = new ResetPassword();

            model.cmpid = EncryptionHelper.Decrypt(tqrs);
            ViewBag.Success = false;
            ViewBag.Error = false;
            return View(model);

        }

        [HttpPost]
        public IActionResult ResetPassword(ResetPassword model)
        {
            if (model != null)
            {
                var user = new tbl_users();

                user.Id = Convert.ToInt32(model.cmpid);
                user.password = EncryptionHelper.Encrypt(model.password);
                _userService.UpdateUserPassword(user);

                ViewBag.Success = true;
                ViewBag.Error = false;
            }
            else
            {
                ViewBag.Success = false;
                ViewBag.Error = true;
            }
            return View();
        }
        [HttpGet]
        public IActionResult ChangePassword(string tqrs)
        {
            tqrs = tqrs.Replace(" ", "+");
            var model = new ChangePassword();

            model.cmpid = EncryptionHelper.Decrypt(tqrs);
            ViewBag.Success = false;
            ViewBag.Error = false;
            return View(model);

        }

        [HttpPost]
        public IActionResult ChangePassword(ChangePassword model)
        {
            if (model != null)
            {
                var user = new tbl_users();

                user.Id = Convert.ToInt32(model.cmpid);
                user.password = EncryptionHelper.Encrypt(model.password);
                user.is_newuser = false;
                _userService.ChnageUserPassword(user);

                ViewBag.Success = true;
                ViewBag.Error = false;
            }
            else
            {
                ViewBag.Success = false;
                ViewBag.Error = true;
            }
            //return View();
            return RedirectToAction("GetProvider");
        }

        [HttpPost]
        public IActionResult UpdateSurgoryCenter(string sProcedureDetailIDs, string sId, string sSCName, string sAssistant, string sSurgeon, DateTime sDate)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var result = _pocservices.UpdatePOCSurgoryCenter(sProcedureDetailIDs, sId, sSCName, sAssistant, sSurgeon, cmpid.Value, sDate);
            return Json(result);
        }

        #region private Method
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
            new LogService().Insert(logdata);
        }
        #endregion 
    }

    public class MyPOCExcelModel
    {
        public string mcode { get; set; }

    }
}