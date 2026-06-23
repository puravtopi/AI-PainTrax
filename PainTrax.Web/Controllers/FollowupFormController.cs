using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Models;
using MS.Services;
using Optivem.Framework.Core.Domain;
using PainTrax.Services;
using PainTrax.Web.AzureServices;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using System.Text.Json;
using System.Text.RegularExpressions;
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class FollowupFormController : Controller
    {
        #region Variables
        private readonly IntakeService service = new IntakeService();
        private readonly LocationsService _locservices = new LocationsService();
        private readonly IWebHostEnvironment _env;
        private readonly AzureAIServices _azureService;
        private readonly ILogger<IntakeFormController> _logger;
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private readonly DiagcodesService _diagcodesService = new DiagcodesService();
        private readonly TreatmentMasterService _treatmentService = new TreatmentMasterService();
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly PatientFUService _fuservices = new PatientFUService();
        private readonly FUPage1Service _fuPage1services = new FUPage1Service();
        private readonly FUOtherService _fuOtherService = new FUOtherService();
        private readonly POCServices _pocservices = new POCServices();
        private readonly Common _commonservices = new Common();
        private readonly UserService _userService = new UserService();
        private readonly SettingsService _settingServices = new SettingsService();
        #endregion


        public FollowupFormController(
        Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
        IWebHostEnvironment env, AzureAIServices azureService,
        ILogger<IntakeFormController> logger
       )
        {
            Environment = environment;
            _env = env;
            _azureService = azureService;
            _logger = logger;

        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Create(int? locId, int? id, int? providerId)
        {
            if (providerId == null)
            {
                if (!HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).HasValue)
                {
                    //providerId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).Value;
                    providerId = Convert.ToInt32(HttpContext.Session.GetString("ProviderId") ?? "0");
                }
                else
                {
                    providerId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).Value;
                }
            }
            var templatePath = $"{Request.Scheme}://{Request.Host}/v2/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            //var templatePath = $"{Request.Scheme}://{Request.Host}/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            ViewBag.TemplateURL = templatePath + "/report-template.txt";
            ViewBag.TemplateDOCURL = templatePath + "/report-template-ie.docx";
            ViewBag.FormData = "";
            ViewBag.Id = "0";
            ViewBag.LocId = locId;

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            tbl_locations objLoc = new tbl_locations()
            {
                id = locId
            };
            var loc = _locservices.GetOne(objLoc);


            ViewBag.LocName = loc?.location;

            var locdata = _commonservices.GetLocations(Convert.ToInt32(cmpid));

            List<SelectListItem> lst = new List<SelectListItem>();

            int defaultlocation = HttpContext.Session.GetInt32(SessionKeys.SessionLocationId).Value;

            foreach (var item in locdata)
            {
                var obj = new SelectListItem()
                {
                    Text = item.Text,
                    Value = item.Value,
                    Selected = item.Value == locId.ToString() ? true : false
                };
                lst.Add(obj);

            }
            ViewBag.locList = lst;
            ViewBag.CmpId = cmpid.ToString();
            var providers = _userService.GetProviders(cmpid.Value);
            List<SelectListItem> lstp = new List<SelectListItem>();
            //int providerid = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).Value;
            foreach (var item in providers)
            {
                var obj = new SelectListItem()
                {
                    Text = item.Text,
                    Value = item.Value,
                    Selected = item.Value == providerId.ToString() ? true : false
                };
                lstp.Add(obj);

            }
            ViewBag.providerList = lstp;

            var _dataTreatment = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);
            ViewBag.Treatment = _dataTreatment;

            if (id != null)
            {
                var data = service.GetInitialIntakeAIById(id.Value);

                if (data != null)
                {
                    ViewBag.FormData = data.FormData;
                    ViewBag.Id = id;
                    ViewBag.SubmitDate = data.PatientSubmitDate;
                    ViewBag.Diagnosis = data.Diagnosis;
                    ViewBag.TreatmentHTML = data.Treatment;

                }
            }

            var client_code = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);

            if (client_code.ToLower() == "qmppc")
                return PartialView("_IntakeQMPPC");
            else if (client_code.ToLower() == "bhfpc")
                return PartialView("_IntakeBHFFU");
            else if (client_code.ToLower() == "hposm")
                return PartialView("_IntakeHPOSM");
            else if (client_code.ToLower() == "imnpfhpc")
                return PartialView("_IntakeIMNPFHPCFU");
            else return PartialView("_IntakeBHFFU");
        }
        [HttpPost]
        public IActionResult Create(FollowupForm model)
        {
            return RedirectToAction("Create", "FollowupForm");
        }

        public IActionResult List(string f = "", string statusFilter = "Active", DateTime? fdate = null, DateTime? tdate = null, int? locid = null)
        {
            try
            {
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

                //ViewBag.locList

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
                string cnd = " and cmp_id = " + cmpid;

                if (statusFilter == "Active")
                {
                    cnd = cnd + " and (is_close=0 or is_close is null) ";
                }
                else if (statusFilter == "Inactive")
                {
                    cnd = cnd + " and (is_close=1) ";
                }
                if (locid == null)
                {
                    locid = HttpContext.Session.GetInt32(SessionKeys.SessionLocationId);
                }
                else
                {
                    locid = locid;
                }


                if (!string.IsNullOrEmpty(searchValue))
                {
                    //cnd = cnd + " and (fname like '%" + searchValue + "%' or lname  like '%" + searchValue + "%' or CONCAT(fname,' ',lname)  LIKE '%" + searchValue + "%' or CONCAT(lname,' ',fname)  LIKE '%" + searchValue + "%' or " +
                    //    "location  like '%" + searchValue + "%' or DATE_FORMAT(dob,\"%m/%d/%Y\") = '" + searchValue + "' or DATE_FORMAT(doe,\"%m/%d/%Y\") = '" + searchValue + "'  or " +
                    //    "compensation like '%" + searchValue + "%' or DATE_FORMAT(doa,\"%m/%d/%Y\") = '" + searchValue + "') ";

                    cnd = cnd + " and ((fname like '%" + searchValue + "%' or lname  like '%" + searchValue + "%' or CONCAT(fname,' ',lname)  LIKE '%" + searchValue + "%' or CONCAT(lname,' ',fname)  LIKE '%" + searchValue + "%' or " +
                      "location  like '%" + searchValue + "%' or DATE_FORMAT(dob,\"%m/%d/%Y\") = '" + searchValue + "' or DATE_FORMAT(doe,\"%m/%d/%Y\") = '" + searchValue + "'  or " +
                      "compensation like '%" + searchValue + "%' or DATE_FORMAT(doa,\"%m/%d/%Y\") = '" + searchValue + "') or " +
                      " id in (SELECT fu.patientIE_ID FROM tbl_patient_fu fu WHERE (DATE_FORMAT(fu.doe,\"%m/%d/%Y\") = '" + searchValue + "')))";


                }
                else
                {
                    if (locid > 0 && (statusFilter == "Active"))
                        cnd = cnd + " and location_id=" + locid;
                }
                if (locid > 0 && string.IsNullOrEmpty(searchValue))
                {
                    cnd = cnd + " and location_id=" + locid;

                }
                //if (fdate != null)
                //{
                //    cnd = cnd  + " (doe = '" + fdate.Value.ToString("yyyy/MM/dd") + "' )";
                //}
                //if (tdate != null)
                //{
                //    cnd = cnd + " (doe = '" + tdate.Value.ToString("yyyy/MM/dd") + "' )";
                //}



                if (!string.IsNullOrEmpty(f))
                {
                    if (f == "A")
                    {
                        cnd = cnd + " AND  attorney_id=0 and patient_id IN(SELECT id FROM tbl_patient WHERE cmp_id = " + cmpid + ")";
                    }
                    else if (f == "I")
                    {
                        cnd = cnd + " AND primary_ins_cmp_id=0 AND patient_id IN (SELECT id FROM tbl_patient WHERE cmp_id=" + cmpid + ")";
                    }
                    else if (f == "C")
                    {
                        cnd = cnd + " AND  (primary_claim_no IS NULL OR primary_claim_no='') AND patient_id IN (SELECT id FROM tbl_patient WHERE cmp_id=" + cmpid + ")";
                    }
                }
                if (fdate != null && tdate != null)
                {
                    cnd += " and DATE(DOE) BETWEEN '"
                           + fdate.Value.ToString("yyyy-MM-dd")
                           + "' AND '"
                           + tdate.Value.ToString("yyyy-MM-dd")
                           + "'";
                }
                var Data = _ieService.GetAll(cnd);


                //Sorting
                if (sortColumn != "0")
                {
                    if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                    {

                        var _sortColumn = Convert.ToInt32(sortColumn);

                        if (_sortColumn > 0)
                            _sortColumn = _sortColumn - 1;

                        var property = typeof(vm_patient_ie).GetProperties()[_sortColumn];
                        if (sortColumnDirection.ToUpper() == "ASC")
                        {
                            Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                        }
                        else
                        {
                            Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                        }
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
                //SaveLog(ex, "List");
            }
            return Json("");

        }

        public JsonResult SearchPatient(string term)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cnd = " and cmp_id = " + cmpid + " and (fname like '%" + term + "%' or lname like '%" + term + "%' or CONCAT(fname, ' ', lname) LIKE CONCAT('%', '" + term + "', '%') or CONCAT(lname, ' ', fname) LIKE CONCAT('%', '" + term + "', '%')) ";
            var result = _ieService.GetAll(cnd);
            return Json(result);
        }

        public IActionResult GetIntakeData(int? id)
        {
            var data = service.GetInitialIntakeFUById(id.Value);
            if (data == null || string.IsNullOrEmpty(data.FormData))
                return Json(new { });

            var formData = string.IsNullOrEmpty(data.FormData)
       ? null
       : System.Text.Json.JsonSerializer.Deserialize<object>(data.FormData);

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            var forwardSetting = _settingServices.GetIntakeOne(cmpid.Value);

            if (forwardSetting == null)
            {
                forwardSetting = new tbl_intake_forward_setting()
                {
                    Neroexam = false,
                    Adl = false,
                    Cc = false,
                    Diagnosis = false,
                    History = false,
                    Note = false,
                    Pe = false
                };
            }

            return Json(new
            {
                data.Id,
                data.Diagnosis,
                FormData = formData,
                ForwardSetting = forwardSetting
            });

        }

        [HttpPost]
        public IActionResult SaveForm([FromBody] object formData)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(formData);
            var model = System.Text.Json.JsonSerializer.Deserialize<AIIntakeFormModel>(json);
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
            var result = "0";
            if (model != null)
            {
                InitialIntakeAI initialIntakeAI = new InitialIntakeAI()
                {
                    Id = model.Id == "" ? 0 : Convert.ToInt32(model.Id),
                    CmpId = cmpid,
                    Visit_Type = "FU",
                    DOA = DateTime.TryParse(model.DOA, out var parsedDOA) ? parsedDOA : (DateTime?)null,
                    DOB = DateTime.TryParse(model.DOB, out var parsedDOB) ? parsedDOB : (DateTime?)null,
                    //DOE = System.DateTime.Now,
                    DOE = DateTime.TryParse(model.DOE, out var parsedDOE) ? parsedDOE : (DateTime?)null,
                    FormData = json,
                    FN = model.FN,
                    LN = model.LN,
                    PatientSubmitDate = DateTime.TryParse(model.PatientSubmitDate, out var PatientSubmitDate) ? parsedDOA : (DateTime?)null,
                    LocationId = string.IsNullOrEmpty(model.LocationId) ? null : Convert.ToInt32(model.LocationId),
                    DLPath = model.DLPath,
                    Diagnosis = model.Diagnosis,
                    Treatment = model.Treatment,
                    TreatmentIds = model.TreatmentIds,
                    TreatmentDelimitDesc = model.TreatmentDelimitDesc
                };
                result = service.SaveInitialIntakeAI(initialIntakeAI);

                var InjuryType = "MM";


                if (model.InjuryType == "work-related")
                    InjuryType = "WC";
                else if (model.InjuryType == "lien")
                    InjuryType = "Lien";
                else
                    InjuryType = model.InjuryType;


                if (initialIntakeAI.Id == 0)
                {

                    tbl_patient_fu objFU = new tbl_patient_fu()
                    {

                        created_by = userid,
                        doe = string.IsNullOrEmpty(model.DOE) ? null : Convert.ToDateTime(model.DOE),
                        patientIE_ID = string.IsNullOrEmpty(model.PatientIEId) ? null : Convert.ToInt32(model.PatientIEId),
                        cmp_id = cmpid,
                        created_date = System.DateTime.Now,
                        is_active = true,
                        patient_id = string.IsNullOrEmpty(model.PatientId) ? null : Convert.ToInt32(model.PatientId),
                        type = InjuryType,
                        intakeid = Convert.ToInt32(result),
                        location_id = string.IsNullOrEmpty(model.LocationId) ? null : Convert.ToInt32(model.LocationId)
                    };

                    var newFU = _fuservices.Insert(objFU);

                    if (newFU > 0)
                    {
                        var objPage1 = new tbl_fu_page1()
                        {
                            pmh = string.Join(", ", model.PMH),
                            psh = string.Join(", ", model.PSH),
                            bodypart = string.Join(",", model.Complaints),
                            allergies = "",
                            assessment = model.Diagnosis,
                            fu_id = newFU,
                            vital = "The patient’s height is " + model.Height + ", weight is " + model.Weight + " pounds, and BMI is _____.",
                            cc = this.GetCC(model)

                        };

                        _fuPage1services.Insert(objPage1);

                        var objOther = new tbl_fu_other()
                        {
                            fu_id = newFU,
                            treatment_delimit = model.TreatmentIds,
                            treatment_delimit_desc = model.TreatmentDelimitDesc,
                            treatment_details = model.TreatmentDesc
                        };
                        _fuOtherService.Insert(objOther);

                        using JsonDocument doc = JsonDocument.Parse(json);

                        string[] planUTPI = doc.RootElement
                                               .GetProperty("PlanUTPI")
                                               .EnumerateArray()
                                               .Select(x => x.GetString())
                                               .ToArray();
                        foreach (string data in planUTPI)
                        {
                            var _obj = new ProcedureDetailsIntakeVM()
                            {
                                PatientIEID = string.IsNullOrEmpty(model.PatientIEId) ? null : Convert.ToInt32(model.PatientIEId),
                                MCode = data,
                                PatientFuID = newFU,
                                Cmp_Id = cmpid.Value,
                                Date = DateTime.TryParse(model.DOE, out var pDOE) ? pDOE : (DateTime?)null,
                                IsExecuted = true
                            };
                            _pocservices.SaveProcedureDetailsIntake(_obj);
                        }

                        string[] recommendation = doc.RootElement
                                               .GetProperty("Recommendation")
                                               .EnumerateArray()
                                               .Select(x => x.GetString())
                                               .ToArray();
                        foreach (string data in recommendation)
                        {
                            var _obj = new ProcedureDetailsIntakeVM()
                            {
                                PatientIEID = string.IsNullOrEmpty(model.PatientIEId) ? null : Convert.ToInt32(model.PatientIEId),
                                MCode = data,
                                PatientFuID = newFU,
                                Cmp_Id = cmpid.Value,
                                Date = DateTime.TryParse(model.DOE, out var pDOE) ? pDOE : (DateTime?)null,
                                IsExecuted = false
                            };
                            _pocservices.SaveProcedureDetailsIntake(_obj);
                        }
                    }
                }
                else
                {
                    var objIE = new tbl_patient_ie()
                    {

                        doa = string.IsNullOrEmpty(model.DOA) ? null : Convert.ToDateTime(model.DOA),
                        doe = string.IsNullOrEmpty(model.DOE) ? null : Convert.ToDateTime(model.DOE),
                        compensation = InjuryType,
                        intakeid = initialIntakeAI.Id
                    };
                    _ieService.UpdateFromIntake(objIE);
                    _ieService.UpdateFromIntakeFU(objIE);
                }

                //return RedirectToAction("Index", "Visit");
            }
            return Json(new { success = true, message = "Intake form summited successfully.", id = result, locid = model.LocationId });
        }

        private string GetCC(AIIntakeFormModel model)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            if (cmpid == "21")
            {
                return "";
            }
            else
            {
                string cc_rsh = "", cc_rsh_difficulty = "", cc_rsh_imporve = "",
                    cc_lsh = "", cc_lsh_difficulty = "", cc_lsh_imporve = "",
                      cc_lkn = "", cc_lkn_difficulty = "", cc_lkn_imporve = "",
                       cc_rkn = "", cc_rkn_difficulty = "", cc_rkn_imporve = "";

                //right soulder
                if (!string.IsNullOrEmpty(model.RShPain))
                    cc_rsh = "The patient’s right shoulder pain level is " + model.RShPain + "/10. ";
                if (model.RShSymptoms?.Count > 0)
                    cc_rsh = cc_rsh + "The patient complains of " + string.Join(", ", model.RShSymptoms) + ". ";

                if (model.RShReachOverhead?.ToLower() == "yes")
                    cc_rsh_difficulty = "Overhead";
                if (model.RShReachBack?.ToLower() == "yes")
                    cc_rsh_difficulty = cc_rsh_difficulty + ", Back";
                if (model.RShSleepIssue?.ToLower() == "yes")
                    cc_rsh_difficulty = cc_rsh_difficulty + ", Sleeping";

                if (!string.IsNullOrEmpty(cc_rsh_difficulty))
                    cc_rsh = cc_rsh + "The patient has difficulty " + cc_rsh_difficulty + " on the right shoulder. ";

                if (model.RShImprove?.Count > 0)
                    cc_rsh = cc_rsh + "There has been improvement with " + string.Join(", ", model.RShImprove) + ".";
                else
                    cc_rsh = cc_rsh + "There has been no improvement with physical therapy.";

                //left soulder
                if (!string.IsNullOrEmpty(model.LShPain))
                    cc_lsh = "The patient’s left shoulder pain level is " + model.LShPain + "/10. ";
                if (model.LShSymptoms?.Count > 0)
                    cc_lsh = cc_lsh + "The patient complains of " + string.Join(", ", model.LShSymptoms) + ". ";


                if (model.LShReachOverhead?.ToLower() == "yes")
                    cc_lsh_difficulty = "Overhead";
                if (model.LShReachBack?.ToLower() == "yes")
                    cc_lsh_difficulty = cc_lsh_difficulty + ", Back";
                if (model.LShSleepIssue?.ToLower() == "yes")
                    cc_lsh_difficulty = cc_lsh_difficulty + ", Sleeping";
                if (!string.IsNullOrEmpty(cc_lsh_difficulty))
                    cc_lsh = cc_lsh + "The patient has difficulty " + cc_lsh_difficulty + " on the left shoulder. ";

                if (model.LShImprove?.Count > 0)
                    cc_lsh = cc_lsh + "There has been improvement with " + string.Join(", ", model.LShImprove) + ".";
                else
                    cc_lsh = cc_lsh + "There has been no improvement with physical therapy.";

                //right knee
                if (!string.IsNullOrEmpty(model.RKnPain))
                    cc_rkn = "The patient’s right knee pain level is " + model.RKnPain + "/10. ";
                if (model.RKnSymptoms?.Count > 0)
                    cc_rkn = cc_rkn + "The patient complains of " + string.Join(", ", model.RKnSymptoms) + ". ";


                if (model.RKnReachOverhead?.ToLower() == "yes")
                    cc_rkn_difficulty = "Overhead";
                if (model.RKnReachBack?.ToLower() == "yes")
                    cc_rkn_difficulty = cc_rkn_difficulty + ", Back";
                if (model.RKnSleepIssue?.ToLower() == "yes")
                    cc_rkn_difficulty = cc_rkn_difficulty + ", Sleeping";

                if (!string.IsNullOrEmpty(cc_rkn_difficulty))
                    cc_rkn = cc_rkn + "The patient has difficulty " + cc_rkn_difficulty.TrimStart(',') + " on the right knee. ";

                if (model.RKnImprove?.Count > 0)
                    cc_rkn = cc_rkn + "There has been improvement with " + string.Join(", ", model.RKnImprove) + ".";
                else
                    cc_rkn = cc_rkn + "There has been no improvement with physical therapy.";

                //left knee
                if (!string.IsNullOrEmpty(model.LKnPain))
                    cc_lkn = "The patient’s left knee pain level is " + model.LKnPain + "/10. ";
                if (model.LKnSymptoms?.Count > 0)
                    cc_lkn = cc_lkn + "The patient complains of " + string.Join(", ", model.LKnSymptoms) + ". ";


                if (model.LKnReachOverhead?.ToLower() == "yes")
                    cc_lkn_difficulty = "Overhead";
                if (model.LKnReachBack?.ToLower() == "yes")
                    cc_lkn_difficulty = cc_lkn_difficulty + ", Back";
                if (model.LKnSleepIssue?.ToLower() == "yes")
                    cc_lkn_difficulty = cc_lkn_difficulty + ", Sleeping";

                if (!string.IsNullOrEmpty(cc_lkn_difficulty))
                    cc_lkn = cc_lkn + "The patient has difficulty " + cc_lkn_difficulty.TrimStart(',') + " on the left knee. ";

                if (model.LKnImprove?.Count > 0)
                    cc_lkn = cc_lkn + "There has been improvement with " + string.Join(", ", model.LKnImprove) + ".";
                else
                    cc_lkn = cc_lkn + "There has been no improvement with physical therapy.";

                return cc_rsh + "<br/>" + cc_lsh + "<br/>" + cc_rkn + "<br/>" + cc_lkn;
            }
        }

        [HttpGet]
        public IActionResult GeneratePdf(string id, string pdffile = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            Dictionary<string, string> controls = new Dictionary<string, string>();



            ParentService _parentService = new ParentService();


            byte[] pdfBytes = null;
            DataTable dt = _parentService.GetData("select * from vm_patient_fu where intakeid=" + id);
            if (dt.Rows.Count > 0)
            {
                PdfHelper _pdfhelper = new PdfHelper();
                string outputfilename = "";
                string fu_id = "";
                var uploadsFolder = "";
                var filePath = "";
                var signPath = "";
                controls.Add("chk_fu", "Yes");
                try
                {
                    DataTable dtdos = _parentService.GetData("select id,doe from tbl_patient_fu  where intakeid=" + id);
                    if (dtdos.Rows.Count > 0)
                    {
                        controls.Add("txt_dos", DateTime.Parse(dtdos.Rows[0]["doe"].ToString()).ToString("MM/dd/yyyy"));
                        fu_id = dtdos.Rows[0]["id"].ToString();
                    }
                }
                catch { }


                try
                {
                    DataTable dtbodypart = _parentService.GetData("select bodypart from tbl_fu_page1  where fu_id=" + fu_id);
                    if (dtbodypart.Rows.Count > 0)
                    {
                        string bodypart = dtbodypart.Rows[0]["bodypart"].ToString().ToLower();
                        string[] bodyparts = bodypart.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                                        .Select(x => x.Trim())
                                                        .ToArray();
                        foreach (string data in bodyparts)
                        {
                            controls.Add(data, "Yes");
                        }
                    }
                }
                catch (Exception ex)
                {
                }

                try
                {
                    DataTable dtplan = _parentService.GetData("select FormData from tbl_intake_ai  where id=" + id);
                    if (dtplan.Rows.Count > 0)
                    {
                        string jsonString = dtplan.Rows[0]["FormData"].ToString().ToLower(); ;

                        using JsonDocument doc = JsonDocument.Parse(jsonString);

                        string[] planUTPI = doc.RootElement
                                               .GetProperty("planutpi")
                                               .EnumerateArray()
                                               .Select(x => x.GetString())
                                               .ToArray();
                        foreach (string data in planUTPI)
                        {
                            if (data.Trim() != "")
                                if (!controls.ContainsKey("utpi"))
                                    controls.Add("utpi", "Yes");
                        }
                    }

                }
                catch (Exception ex)
                {

                }


                try
                {
                    uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Downloads/" + cmpclientid);
                    filePath = Path.Combine(uploadsFolder, pdffile);

                    //  signPath = Path.Combine(Directory.GetCurrentDirectory(), "signatures");
                }
                catch (Exception ex)
                {
                    //SaveLog(ex, "set Paths");
                }


                try
                {
                    pdfBytes = _pdfhelper.Stamping(filePath, "Id", dt.Rows[0]["patient_id"].ToString(), controls, cmpid, signPath);
                }
                catch (Exception ex)
                {
                    //SaveLog(ex, "Pdf Stamping");
                }
                string fileName = $"Superbill.pdf";
                string folder = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments/Others/" + dt.Rows[0]["patient_id"].ToString());

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                string destfilePath = Path.Combine(folder, fileName);

                System.IO.File.WriteAllBytes(destfilePath, pdfBytes);

                //string htmlContent = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "demo.html"));
                // ViewBag.FileName = dt.Rows[0]["LastName"].ToString() + " " + dt.Rows[0]["FirstName"].ToString();


            }

            return File(pdfBytes, "application/pdf", $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Superbill.pdf");
        }
    }
}
