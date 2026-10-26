using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Models;
using MS.Services;
using MySqlX.XDevAPI;
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
        private readonly ForwardServices _forwardServices = new ForwardServices();
        private readonly DefaultDataServices _defaultService = new DefaultDataServices();
        private readonly FUPage3Service _fuPage3services = new FUPage3Service();
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
        public IActionResult Create(int? locId, int? id, int? providerId, int patientId = 0, string fType = "FU",
            int ie_id = 0, int fu_id = 0)
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
            ViewBag.ProviderId = providerId;
            ViewBag.IEId = ie_id;
            ViewBag.FUId = fu_id;

            ViewBag.Type = (fType == "TelFU" || fType == "Tele FU") ? "Tele FU" : "FU";

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
            ViewBag.ATList = _commonservices.GetAccidenttype(cmpid.Value);
            var client_code = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            ViewBag.PatientId = patientId;
            if (client_code.ToLower() == "qmppc")
                return PartialView("_IntakeQMPPC");
            else if (client_code.ToLower() == "bhfpc")
                return PartialView("_IntakeBHFFU");
            else if (client_code.ToLower() == "paintest")
                return PartialView("_IntakeKDVPCFU");
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

            if (formData is System.Text.Json.JsonElement element)
            {
                // Check if property exists and get it
                if (element.TryGetProperty("FN", out var property))
                {
                    string value = property.GetString();
                    // or property.GetRawText() / property.GetInt32()
                }
            }

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
            json = json.Replace("\\u0027", "'");
            var model = new AIIntakeFormModel();
            try
            {
                model = System.Text.Json.JsonSerializer.Deserialize<AIIntakeFormModel>(json);
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Path: {ex.Path}");
                Console.WriteLine(ex.Message);
            }
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

            var result = "0";
            int ieId = 0, newFU = 0;
            if (model != null)
            {

                InitialIntakeAI initialIntakeAI = new InitialIntakeAI()
                {
                    Id = model.Id == "" ? 0 : Convert.ToInt32(model.Id),
                    PatientIEId = model.PatientIEId == "" ? 0 : Convert.ToInt32(model.PatientIEId),
                    CmpId = cmpid,
                    Visit_Type = model.VisitType,
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

                ieId = initialIntakeAI.PatientIEId.Value;



                string ccContent = "", peContent = "", dtrContent = "", sensoryContent = "", mmstContent = "";
               
                if (HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToLower() == "bhfpc" ||
                    HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToLower() == "ipmc")
                {
                    peContent = this.GetPE(model);
                    dtrContent = this.GetDTR(model);
                    sensoryContent = this.GetSensory(model);
                    mmstContent = this.GetMMST(model);
                    ccContent = this.GetBHFCC(model);
                }
                



                if (initialIntakeAI.Id == 0)
                {
                    var data = _ieService.CheckFUDOS(initialIntakeAI.PatientIEId.Value, initialIntakeAI.DOE);

                    if (data)
                    {
                        return Json(new { success = false, message = "A visit has already been recorded for this patient on the selected DOS.", id = 0, locid = model.LocationId, provid = model.ProviderId });
                    }
                }


                result = service.SaveInitialIntakeAI(initialIntakeAI);

                var InjuryType = "MM";
                var Handeness = "1";
                var patientId = 0;

                if (model.DominantHand != null)
                {
                    if (model.DominantHand.ToLower() == "right-handed")
                        Handeness = "1";
                    else if (model.DominantHand.ToLower() == "left-handed")
                        Handeness = "2";
                }

                if (model.InjuryType == "work-related")
                    InjuryType = "WC";
                else if (model.InjuryType == "lien")
                    InjuryType = "Lien";
                else
                    InjuryType = model.InjuryType;

                int lFUId = 0;

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
                        type = model.VisitType,
                        intakeid = Convert.ToInt32(result),
                        location_id = string.IsNullOrEmpty(model.LocationId) ? null : Convert.ToInt32(model.LocationId),
                        provider_id = string.IsNullOrEmpty(model.ProviderId) ? null : Convert.ToInt32(model.ProviderId),
                        accident_type = model.AccidentType,
                    };

                    var fuData = _ieService.GetLastFU(objFU.patientIE_ID.Value, "FU");

                    if (fuData != null)
                    {
                        if (fuData.procedure_performed != null)
                        {
                            objFU.procedure_performed = fuData.procedure_performed;
                        }
                    }
                    newFU = _fuservices.Insert(objFU);

                    if (newFU > 0)
                    {
                        //var objPage1 = new tbl_fu_page1()
                        //{
                        //    pmh = model.PMH != null ? string.Join(", ", model.PMH) : "",
                        //    psh = model.PSH != null ? string.Join(", ", model.PSH) : "",
                        //    bodypart = model.Complaints != null ? string.Join(",", model.Complaints) : "",
                        //    allergies = "",
                        //    assessment = model.Diagnosis,
                        //    fu_id = newFU,
                        //    vital = "The patient’s height is " + model.Height + ", weight is " + model.Weight + " pounds, and BMI is _____.",
                        //    cc = this.GetCC(model)

                        //};

                        //_fuPage1services.Insert(objPage1);

                        //var objOther = new tbl_fu_other()
                        //{
                        //    fu_id = newFU,
                        //    treatment_delimit = model.TreatmentIds,
                        //    treatment_delimit_desc = model.TreatmentDelimitDesc,
                        //    treatment_details = model.TreatmentDesc
                        //};
                        //_fuOtherService.Insert(objOther);


                        if (fuData != null)
                            lFUId = fuData.id.Value;


                        try
                        {
                            _forwardServices.GetOnePage1(objFU.patientIE_ID.Value, newFU, cmpid.Value, objFU.patient_id.Value, lFUId);
                        }
                        catch (Exception ex)
                        {
                        }

                        _ieService.UpdateBodyPartFromIntakeFU(string.Join(",", model.Complaints), newFU);

                        try
                        {
                            _forwardServices.GetOnePage2(objFU.patientIE_ID.Value, newFU, cmpid.Value, objFU.patient_id.Value, lFUId);
                        }
                        catch (Exception ex)
                        {
                        }

                        try
                        {
                            _forwardServices.GetOneNE(objFU.patientIE_ID.Value, newFU, cmpid.Value, objFU.patient_id.Value, lFUId);
                        }
                        catch (Exception ex)
                        {
                        }

                        try
                        {
                            var follwup = HttpContext.Session.GetString(SessionKeys.SessionFUDate);
                            _forwardServices.GetOneOther(objFU.patientIE_ID.Value, newFU, cmpid.Value, objFU.patient_id.Value, lFUId, follwup);
                        }
                        catch (Exception ex)
                        {
                        }
                        try
                        {
                            _forwardServices.GetPage3(objFU.patientIE_ID.Value, cmpid.Value, newFU, objFU.patient_id.Value, lFUId);
                        }
                        catch (Exception ex)
                        {
                        }


                        using JsonDocument doc = JsonDocument.Parse(json);

                        string[] planUTPI = Array.Empty<string>();
                        if (doc.RootElement.TryGetProperty("PlanUTPI", out JsonElement planElement))
                        {
                            // 2. Optional: Ensure it is actually an array before enumerating
                            if (planElement.ValueKind == JsonValueKind.Array)
                            {
                                planUTPI = planElement
                                    .EnumerateArray()
                                    .Select(x => x.GetString() ?? string.Empty)
                                    .ToArray();
                            }
                        }
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
                        string[] recommendation = Array.Empty<string>();
                        if (doc.RootElement.TryGetProperty("Recommendation", out JsonElement recommendationElement))
                        {

                            // 2. Optional: Ensure it is actually an array before enumerating
                            if (recommendationElement.ValueKind == JsonValueKind.Array)
                            {
                                recommendation = recommendationElement
                                    .EnumerateArray()
                                    .Select(x => x.GetString() ?? string.Empty)
                                    .ToArray();
                            }
                        }
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
                        intakeid = initialIntakeAI.Id,
                        id = string.IsNullOrEmpty(model.PatientIEId) ? null : Convert.ToInt32(model.PatientIEId),
                        provider_id = string.IsNullOrEmpty(model.ProviderId) ? null : Convert.ToInt32(model.ProviderId),
                    };
                    var fuData = _ieService.GetLastFU(objIE.id.Value, "FU");
                    ieId = objIE.id.Value;


                    if (fuData != null)
                        lFUId = fuData.id.Value;
                    // _ieService.UpdateFromIntake(objIE);
                    _ieService.UpdateFromIntakeFU(objIE);
                    _ieService.UpdateCompensationFU(objIE);
                    _ieService.UpdateBodyPartFromIntakeFU(string.Join(",", model.Complaints), lFUId);
                    newFU = lFUId;
                }
                patientId = string.IsNullOrEmpty(model.PatientId) ? 0 : Convert.ToInt32(model.PatientId);
                _ieService.UpdateHandeness(Handeness, patientId);
                //return RedirectToAction("Index", "Visit");
            }
            return Json(new { success = true, message = "Intake form summited successfully.", id = result, locid = model.LocationId, provid = model.ProviderId, patientid = model.PatientId, ie_id = ieId, fu_id = newFU,fu_type= model.VisitType });
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
                    controls.Add("location", dt.Rows[0]["location"].ToString());
                }
                catch { }

                try
                {
                    DataTable dtdos = _parentService.GetData("select id,doe,provider_id from tbl_patient_fu  where intakeid=" + id);
                    if (dtdos.Rows.Count > 0)
                    {
                        controls.Add("txt_dos", DateTime.Parse(dtdos.Rows[0]["doe"].ToString()).ToString("MM/dd/yyyy"));
                        fu_id = dtdos.Rows[0]["id"].ToString();
                        DataTable dtpro = _parentService.GetData("select id,fullname,desigid,desig_name from vm_cm_user  where id=" + dtdos.Rows[0]["provider_id"].ToString());
                        if (dtpro.Rows.Count > 0)
                        {
                            if (dtpro.Rows[0]["fullname"].ToString().StartsWith("Irina"))
                            {
                                controls.Add("chk_irina", "Yes");
                            }
                            else
                            {
                                controls.Add("chk_other", "Yes");
                                controls.Add("txt_provname", dtpro.Rows[0]["fullname"].ToString());
                            }


                        }
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
                        try
                        {
                            string[] planUTPI = doc.RootElement
                                                   .GetProperty("planutpi")
                                                   .EnumerateArray()
                                                   .Select(x => x.GetString())
                                                   .ToArray();
                            foreach (string data in planUTPI)
                            {
                                if (data.Trim().ToLower() == "ctpi" || data.Trim().ToLower() == "ttpi" || data.Trim().ToLower() == "ltpi")
                                    if (!controls.ContainsKey("utpi"))
                                        controls.Add("utpi", "Yes");
                                if (data.Trim().ToLower() == "pt")
                                    if (!controls.ContainsKey("pt"))
                                        controls.Add("pt", "Yes");

                            }
                        }
                        catch { }

                        try
                        {
                            string[] planUTPI = doc.RootElement
                                                   .GetProperty("planmri")
                                                   .EnumerateArray()
                                                   .Select(x => x.GetString())
                                                   .ToArray();
                            foreach (string data in planUTPI)
                            {
                                if (data.Trim().ToLower() == "cspine")
                                    if (!controls.ContainsKey("cspine"))
                                        controls.Add("cspine", "Yes");
                                if (data.Trim().ToLower() == "tspine")
                                    if (!controls.ContainsKey("tspine"))
                                        controls.Add("tspine", "Yes");
                                if (data.Trim().ToLower() == "lspine")
                                    if (!controls.ContainsKey("lspine"))
                                        controls.Add("lspine", "Yes");

                            }
                        }
                        catch { }

                        try
                        {
                            string[] planUTPI = doc.RootElement
                                                   .GetProperty("planxray")
                                                   .EnumerateArray()
                                                   .Select(x => x.GetString())
                                                   .ToArray();
                            foreach (string data in planUTPI)
                            {
                                if (data.Trim().ToLower() == "cspine")
                                    if (!controls.ContainsKey("cspine"))
                                        controls.Add("cspine", "Yes");
                                if (data.Trim().ToLower() == "tspine")
                                    if (!controls.ContainsKey("tspine"))
                                        controls.Add("tspine", "Yes");
                                if (data.Trim().ToLower() == "lspine")
                                    if (!controls.ContainsKey("lspine"))
                                        controls.Add("lspine", "Yes");

                            }
                        }
                        catch { }

                        try
                        {
                            string[] planUTPI = doc.RootElement
                                                   .GetProperty("planct")
                                                   .EnumerateArray()
                                                   .Select(x => x.GetString())
                                                   .ToArray();
                            foreach (string data in planUTPI)
                            {
                                if (data.Trim().ToLower() == "cspine")
                                    if (!controls.ContainsKey("cspine"))
                                        controls.Add("cspine", "Yes");
                                if (data.Trim().ToLower() == "tspine")
                                    if (!controls.ContainsKey("tspine"))
                                        controls.Add("tspine", "Yes");
                                if (data.Trim().ToLower() == "lspine")
                                    if (!controls.ContainsKey("lspine"))
                                        controls.Add("lspine", "Yes");

                            }
                        }
                        catch { }
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
                string fileName = $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Superbill_{DateTime.Parse(dt.Rows[0]["doe"].ToString()).ToString("MMddyyyy")}.pdf";
                string folder = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments/Other/" + dt.Rows[0]["patient_id"].ToString());

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                string destfilePath = Path.Combine(folder, fileName);

                System.IO.File.WriteAllBytes(destfilePath, pdfBytes);

                //string htmlContent = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "demo.html"));
                // ViewBag.FileName = dt.Rows[0]["LastName"].ToString() + " " + dt.Rows[0]["FirstName"].ToString();


            }

            return File(pdfBytes, "application/pdf", $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Superbill_{DateTime.Parse(dt.Rows[0]["doe"].ToString()).ToString("MMddyyyy")}.pdf");
        }


        [HttpGet]
        public IActionResult CheckSign(string id)
        {
            ParentService _parentService = new ParentService();
            DataTable dt = _parentService.GetData("select * from vm_patient_fu where intakeid=" + id);
            if (dt.Rows.Count > 0)
            {
                string pid = dt.Rows[0]["patient_id"].ToString();
                var data = new tbl_ie_sign();
                try
                {
                    data = _ieService.GetOnesign(Convert.ToInt32(pid));
                    if (data == null)
                        return Json(new { id = 0 });
                    return Json(new { id = data.patient_id }); ;
                }
                catch (Exception ex)
                {
                    return Json(new { id = 0 });
                }

            }
            return Json(new { id = 0 });

        }


        [HttpGet]
        public IActionResult GenerateCF(string id, string pdffile = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            Dictionary<string, string> controls = new Dictionary<string, string>();
            int provid = 0;


            ParentService _parentService = new ParentService();


            byte[] pdfBytes = null;
            DataTable dt = _parentService.GetData("select * from vm_patient_fu where intakeid=" + id);
            if (dt.Rows.Count > 0)
            {
                PdfHelper _pdfhelper = new PdfHelper();
                string outputfilename = "";
                string ie_id = "";
                var uploadsFolder = "";
                var filePath = "";
                var signPath = "";

                try
                {
                    DataTable dtdos = _parentService.GetData("select id,doe,provider_id from tbl_patient_fu where intakeid=" + id);
                    if (dtdos.Rows.Count > 0)
                    {
                        controls.Add("txt_date", DateTime.Parse(dtdos.Rows[0]["doe"].ToString()).ToString("MM/dd/yyyy"));
                        ie_id = dtdos.Rows[0]["id"].ToString();
                        DataTable dtpro = _parentService.GetData("select id,fullname,desigid,desig_name from vm_cm_user  where id=" + dtdos.Rows[0]["provider_id"].ToString());
                        if (dtpro.Rows.Count > 0)
                        {

                            controls.Add("txt_provname", dtpro.Rows[0]["fullname"].ToString());
                            provid = Convert.ToInt32(dtdos.Rows[0]["provider_id"].ToString());
                        }
                    }
                }
                catch { }

                try
                {
                    DataTable dtplan = _parentService.GetData("select FormData from tbl_intake_ai where id=" + id);

                    if (dtplan.Rows.Count > 0)
                    {
                        string jsonString = dtplan.Rows[0]["FormData"].ToString().ToLower();

                        using JsonDocument doc = JsonDocument.Parse(jsonString);

                        string[] planUTPI = doc.RootElement
                                               .GetProperty("planutpi")
                                               .EnumerateArray()
                                               .Select(x => x.GetString()?.Trim().ToLower())
                                               .ToArray();

                        List<string> tpiList = new List<string>();

                        foreach (string data in planUTPI)
                        {
                            switch (data)
                            {
                                case "ctpi":
                                    tpiList.Add("Cervical");
                                    break;

                                case "ttpi":
                                    tpiList.Add("Thoracic");
                                    break;

                                case "ltpi":
                                    tpiList.Add("Lumbar");
                                    break;


                            }
                        }

                        if (tpiList.Count > 0)
                        {
                            controls["txt_tpi"] = string.Join(", ", tpiList) + " TPI";
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

                    signPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/signatures/");
                }
                catch (Exception ex)
                {
                    //SaveLog(ex, "set Paths");
                }


                try
                {
                    pdfBytes = _pdfhelper.Stamping(filePath, "Id", dt.Rows[0]["patient_id"].ToString(), controls, cmpid, signPath, provid);
                }
                catch (Exception ex)
                {
                    // SaveLog(ex, "Pdf Stamping");
                }
                string fileName = $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Consent Form_{DateTime.Parse(dt.Rows[0]["doe"].ToString()).ToString("MMddyyyy")}.pdf";
                string folder = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments/Other/" + dt.Rows[0]["patient_id"].ToString());

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                string destfilePath = Path.Combine(folder, fileName);

                System.IO.File.WriteAllBytes(destfilePath, pdfBytes);

                //string htmlContent = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "demo.html"));
                // ViewBag.FileName = dt.Rows[0]["LastName"].ToString() + " " + dt.Rows[0]["FirstName"].ToString();


            }

            return File(pdfBytes, "application/pdf", $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Consent Form_{DateTime.Parse(dt.Rows[0]["doe"].ToString()).ToString("MMddyyyy")}.pdf");
        }

        #region AI Content
        private string GetBHFCC(AIIntakeFormModel model)
        {
            string pe = "";
            if (model.Complaints.Contains("neck", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + CCHelper.NeckCC(model);
            }
            if (model.Complaints.Contains("midback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.MidbackCC(model) + "</p>";
            }
            if (model.Complaints.Contains("lowback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LowBackCC(model) + "</p>";
            }
            if (model.Complaints.Contains("right shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.RightShoulderCC(model) + "</p>";
            }
            if (model.Complaints.Contains("left shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LeftShoulderCC(model) + "</p>";
            }
            if (model.Complaints.Contains("right knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.RightKneeCC(model) + "</p>";
            }
            if (model.Complaints.Contains("left knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LeftKneeCC(model) + "</p>";
            }
            if (model.Complaints.Contains("right ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.RightAnkleCC(model) + "</p>";
            }
            if (model.Complaints.Contains("left ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LeftAnkleCC(model) + "</p>";
            }
            if (model.Complaints.Contains("right Elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.RightElbowCC(model) + "</p>";
            }
            if (model.Complaints.Contains("left elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LeftElbowCC(model) + "</p>";
            }
            if (model.Complaints.Contains("right hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.RightHipCC(model) + "</p>";
            }
            if (model.Complaints.Contains("left hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LeftHipCC(model) + "</p>";
            }
            if (model.Complaints.Contains("right wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.RightWristCC(model) + "</p>";
            }
            if (model.Complaints.Contains("left wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/><p>" + CCHelper.LeftWristCC(model) + "</p>";
            }
            return pe.Replace(", and", " and");
        }

        private string GetPE(AIIntakeFormModel model)
        {
            string pe = "";

            if (model.Complaints.Contains("neck", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "neck", "Neck.html");
            }
            if (model.Complaints.Contains("midback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "midback", "Midback.html");
            }
            if (model.Complaints.Contains("lowback", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "lowback", "Lowback.html");
            }
            if (model.Complaints.Contains("right shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "right shoulder", "Shoulder.html");
            }
            if (model.Complaints.Contains("left shoulder", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "left shoulder", "Shoulder.html");
            }
            if (model.Complaints.Contains("right knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "right knee", "Knee.html");
            }
            if (model.Complaints.Contains("left knee", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "left knee", "Knee.html");
            }
            if (model.Complaints.Contains("right ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "right ankle", "Ankle.html");
            }
            if (model.Complaints.Contains("left ankle", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "left ankle", "Ankle.html");
            }
            if (model.Complaints.Contains("right Elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "right elbow", "Elbow.html");
            }
            if (model.Complaints.Contains("left elbow", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "left elbow", "Elbow.html");
            }
            if (model.Complaints.Contains("right hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "right hip", "Hip.html");
            }
            if (model.Complaints.Contains("left hip", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "left hip", "Hip.html");
            }
            if (model.Complaints.Contains("right wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "right wrist", "Wrist.html");
            }
            if (model.Complaints.Contains("left wrist", StringComparer.OrdinalIgnoreCase))
            {
                pe = pe + "<br/>" + this.GetROM(model, "left wrist", "Wrist.html");
            }

            return pe;
        }

        private string GetDTR(AIIntakeFormModel model)
        {
            try
            {

                var templateDir = Path.Combine(Environment.WebRootPath, "templates");

                //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
                var filePath = templateDir + "/BHF/DTR.html";


                // 2. Read the entire file content into a string asynchronously
                string htmlString = System.IO.File.ReadAllText(filePath);

                // 2. Parse HTML
                /*var parser = new HtmlParser();
                var document = parser.ParseDocument(htmlString);

                string UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
                string LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

                string className = "";

                bool isUEPresent = model.Complaints.Any(value =>
                {
                    string sBody = value.Replace("Left", "")
                                        .Replace("Right", "")
                                        .Trim();

                    // In C#, you can do a case-insensitive check directly without calling ToLower()
                    return UEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
                });

                if (!isUEPresent)
                {
                    // 3. Select all elements matching the class and remove them
                    className = "upper"; // Target class
                    var elements = document.QuerySelectorAll($".{className}");

                    foreach (var element in elements)
                    {
                        element.Remove(); // Completely removes the element
                    }
                }

                bool isLEPresent = model.Complaints.Any(value =>
                {
                    string sBody = value.Replace("Left", "")
                                        .Replace("Right", "")
                                        .Trim();

                    // In C#, you can do a case-insensitive check directly without calling ToLower()
                    return LEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
                });

                if (!isLEPresent)
                {
                    // 3. Select all elements matching the class and remove them
                    className = "lower"; // Target class
                    var elements = document.QuerySelectorAll($".{className}");

                    foreach (var element in elements)
                    {
                        element.Remove(); // Completely removes the element
                    }
                }

                if (!isUEPresent && !isLEPresent)
                {  // 3. Select all elements matching the class and remove them
                    className = "hidAll"; // Target class
                    var elements = document.QuerySelectorAll($".{className}");

                    foreach (var element in elements)
                    {
                        element.Remove(); // Completely removes the element
                    }

                }*/

                // 4. Output the result
                //string modifiedHtml = document.Body.InnerHtml; // or document.DocumentElement.OuterHtml
                string modifiedHtml = htmlString; // or document.DocumentElement.OuterHtml

                modifiedHtml = modifiedHtml.Replace("#TricepsLeft", string.IsNullOrEmpty(model.TricepsLeft) ? "2+" : model.TricepsLeft)
                                              .Replace("#TricepsRight", string.IsNullOrEmpty(model.TricepsRight) ? "2+" : model.TricepsRight)
                                              .Replace("#BicepsLeft", string.IsNullOrEmpty(model.BicepsLeft) ? "2+" : model.BicepsLeft)
                                              .Replace("#BicepsRight", string.IsNullOrEmpty(model.BicepsRight) ? "2+" : model.BicepsRight)
                                              .Replace("#BrachioradialisLeft", string.IsNullOrEmpty(model.BrachioradialisLeft) ? "2+" : model.BrachioradialisLeft)
                                              .Replace("#BrachioradialisRight", string.IsNullOrEmpty(model.BrachioradialisRight) ? "2+" : model.BrachioradialisRight)
                                              .Replace("#KneeRELeft", string.IsNullOrEmpty(model.KneeRELeft) ? "2+" : model.KneeRELeft)
                                              .Replace("#KneeRERight", string.IsNullOrEmpty(model.KneeRERight) ? "2+" : model.KneeRERight)
                                              .Replace("#AnkleRELeft", string.IsNullOrEmpty(model.AnkleRELeft) ? "2+" : model.AnkleRELeft)
                                              .Replace("#AnkleRERight", string.IsNullOrEmpty(model.AnkleRERight) ? "2+" : model.AnkleRERight);
                return modifiedHtml;
            }
            catch (Exception ex)
            {
               // this.SaveLog(ex, "");
                return "";
            }

        }

        private string GetSensory(AIIntakeFormModel model)
        {

            var templateDir = Path.Combine(Environment.WebRootPath, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/Sensory.html";


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);

            // 2. Parse HTML
            //var parser = new HtmlParser();
            //var document = parser.ParseDocument(htmlString);

            //string UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
            //string LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

            //string className = "";

            //bool isUEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return UEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isUEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "upper"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //bool isLEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return LEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isLEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "lower"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //if (!isUEPresent && !isLEPresent)
            //{  // 3. Select all elements matching the class and remove them
            //    className = "hidAll"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }

            //}

            //// 4. Output the result
            //string modifiedHtml = document.Body.InnerHtml;

            htmlString = htmlString.Replace("#UpperShoulderLeft", string.IsNullOrEmpty(model.UpperShoulderLeft) ? "Normal" : model.UpperShoulderLeft)
                                          .Replace("#UpperShoulderRight", string.IsNullOrEmpty(model.UpperShoulderRight) ? "Normal" : model.UpperShoulderRight)
                                          .Replace("#LateralArmLeft", string.IsNullOrEmpty(model.LateralArmLeft) ? "Normal" : model.LateralArmLeft)
                                          .Replace("#LateralArmRight", string.IsNullOrEmpty(model.LateralArmRight) ? "Normal" : model.LateralArmRight)
                                          .Replace("#ForearmC6Left", string.IsNullOrEmpty(model.ForearmC6Left) ? "Normal" : model.ForearmC6Left)
                                          .Replace("#ForearmC6Right", string.IsNullOrEmpty(model.ForearmC6Right) ? "Normal" : model.ForearmC6Right)
                                          .Replace("#MiddleFingerLeft", string.IsNullOrEmpty(model.MiddleFingerLeft) ? "Normal" : model.MiddleFingerLeft)
                                          .Replace("#MiddleFingerRight", string.IsNullOrEmpty(model.MiddleFingerRight) ? "Normal" : model.MiddleFingerRight)
                                          .Replace("#ForearmC8Left", string.IsNullOrEmpty(model.ForearmC8Left) ? "Normal" : model.ForearmC8Left)
                                          .Replace("#ForearmC8Right", string.IsNullOrEmpty(model.ForearmC8Right) ? "Normal" : model.ForearmC8Right)
                                          .Replace("#MedialArmLeft", string.IsNullOrEmpty(model.MedialArmLeft) ? "Normal" : model.MedialArmLeft)
                                          .Replace("#MedialArmRight", string.IsNullOrEmpty(model.MedialArmRight) ? "Normal" : model.MedialArmRight)
                                          .Replace("#CervicalLeft", string.IsNullOrEmpty(model.CervicalLeft) ? "Normal" : model.CervicalLeft)
                                          .Replace("#CervicalRight", string.IsNullOrEmpty(model.CervicalRight) ? "Normal" : model.CervicalRight)
                                          .Replace("#ThighLeft", string.IsNullOrEmpty(model.ThighLeft) ? "Normal" : model.ThighLeft)
                                          .Replace("#ThighRight", string.IsNullOrEmpty(model.ThighRight) ? "Normal" : model.ThighRight)
                                          .Replace("#FootMedialLeft", string.IsNullOrEmpty(model.FootMedialLeft) ? "Normal" : model.FootMedialLeft)
                                          .Replace("#FootMedialRight", string.IsNullOrEmpty(model.FootMedialRight) ? "Normal" : model.FootMedialRight)
                                          .Replace("#FootDorsumLeft", string.IsNullOrEmpty(model.FootDorsumLeft) ? "Normal" : model.FootDorsumLeft)
                                          .Replace("#FootDorsumRight", string.IsNullOrEmpty(model.FootDorsumRight) ? "Normal" : model.FootDorsumRight)
                                          .Replace("#FootLateralLeft", string.IsNullOrEmpty(model.FootLateralLeft) ? "Normal" : model.FootLateralLeft)
                                          .Replace("#FootLateralRight", string.IsNullOrEmpty(model.FootLateralRight) ? "Normal" : model.FootLateralRight)
                                          .Replace("#LumbarLeft", string.IsNullOrEmpty(model.LumbarLeft) ? "Normal" : model.LumbarLeft)
                                          .Replace("#LumbarRight", string.IsNullOrEmpty(model.LumbarRight) ? "Normal" : model.LumbarRight);
            return htmlString;

        }

        private string GetMMST(AIIntakeFormModel model)
        {

            var templateDir = Path.Combine(Environment.WebRootPath, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/MMST.html";


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);

            //// 2. Parse HTML
            //var parser = new HtmlParser();
            //var document = parser.ParseDocument(htmlString);

            //string UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
            //string LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

            //string className = "";

            //bool isUEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return UEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isUEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "upper"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //bool isLEPresent = model.Complaints.Any(value =>
            //{
            //    string sBody = value.Replace("Left", "")
            //                        .Replace("Right", "")
            //                        .Trim();

            //    // In C#, you can do a case-insensitive check directly without calling ToLower()
            //    return LEBodyPart.Contains(sBody, StringComparison.OrdinalIgnoreCase);
            //});

            //if (!isLEPresent)
            //{
            //    // 3. Select all elements matching the class and remove them
            //    className = "lower"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }
            //}

            //if (!isUEPresent && !isLEPresent)
            //{  // 3. Select all elements matching the class and remove them
            //    className = "hidAll"; // Target class
            //    var elements = document.QuerySelectorAll($".{className}");

            //    foreach (var element in elements)
            //    {
            //        element.Remove(); // Completely removes the element
            //    }

            //}

            //// 4. Output the result
            //string modifiedHtml = document.Body.InnerHtml;

            htmlString = htmlString.Replace("#ShoulderAbductionLeft", string.IsNullOrEmpty(model.ShoulderAbductionLeft) ? "5" : model.ShoulderAbductionLeft)
                                          .Replace("#ShoulderAbductionRight", string.IsNullOrEmpty(model.ShoulderAbductionRight) ? "5" : model.ShoulderAbductionRight)
                                          .Replace("#ShoulderFlexionLeft", string.IsNullOrEmpty(model.ShoulderFlexionLeft) ? "5" : model.ShoulderFlexionLeft)
                                          .Replace("#ShoulderFlexionRight", string.IsNullOrEmpty(model.ShoulderFlexionRight) ? "5" : model.ShoulderFlexionRight)
                                          .Replace("#ElbowExtensionLeft", string.IsNullOrEmpty(model.ElbowExtensionLeft) ? "5" : model.ElbowExtensionLeft)
                                          .Replace("#ElbowExtensionRight", string.IsNullOrEmpty(model.ElbowExtensionRight) ? "5" : model.ElbowExtensionRight)
                                          .Replace("#ElbowFlexionLeft", string.IsNullOrEmpty(model.ElbowFlexionLeft) ? "5" : model.ElbowFlexionLeft)
                                          .Replace("#ElbowFlexionRight", string.IsNullOrEmpty(model.ElbowFlexionRight) ? "5" : model.ElbowFlexionRight)
                                          .Replace("#ElbowSupinationLeft", string.IsNullOrEmpty(model.ElbowSupinationLeft) ? "5" : model.ElbowSupinationLeft)
                                          .Replace("#ElbowSupinationRight", string.IsNullOrEmpty(model.ElbowSupinationRight) ? "5" : model.ElbowSupinationRight)
                                          .Replace("#ElbowPronationLeft", string.IsNullOrEmpty(model.ElbowPronationLeft) ? "5" : model.ElbowPronationLeft)
                                          .Replace("#ElbowPronationRight", string.IsNullOrEmpty(model.ElbowPronationRight) ? "5" : model.ElbowPronationRight)
                                          .Replace("#WristFlexionLeft", string.IsNullOrEmpty(model.WristFlexionLeft) ? "5" : model.WristFlexionLeft)
                                          .Replace("#WristFlexionRight", string.IsNullOrEmpty(model.WristFlexionRight) ? "5" : model.WristFlexionRight)
                                          .Replace("#WristExtensionLeft", string.IsNullOrEmpty(model.WristExtensionLeft) ? "5" : model.WristExtensionLeft)
                                          .Replace("#WristExtensionRight", string.IsNullOrEmpty(model.WristExtensionRight) ? "5" : model.WristExtensionRight)
                                          .Replace("#GripLeft", string.IsNullOrEmpty(model.GripLeft) ? "5" : model.GripLeft)
                                          .Replace("#GripRight", string.IsNullOrEmpty(model.GripRight) ? "5" : model.GripRight)
                                          .Replace("#FingerAbductionLeft", string.IsNullOrEmpty(model.FingerAbductionLeft) ? "5" : model.FingerAbductionLeft)
                                          .Replace("#FingerAbductionRight", string.IsNullOrEmpty(model.FingerAbductionRight) ? "5" : model.FingerAbductionRight)
                                          .Replace("#HipFlexionLeft", string.IsNullOrEmpty(model.HipFlexionLeft) ? "5" : model.HipFlexionLeft)
                                          .Replace("#HipFlexionRight", string.IsNullOrEmpty(model.HipFlexionRight) ? "5" : model.HipFlexionRight)
                                          .Replace("#HipAbductionLeft", string.IsNullOrEmpty(model.HipAbductionLeft) ? "5" : model.HipAbductionLeft)
                                          .Replace("#HipAbductionRight", string.IsNullOrEmpty(model.HipAbductionRight) ? "5" : model.HipAbductionRight)
                                          .Replace("#KneeExtensionLeft", string.IsNullOrEmpty(model.KneeExtensionLeft) ? "5" : model.KneeExtensionLeft)
                                          .Replace("#KneeExtensionRight", string.IsNullOrEmpty(model.KneeExtensionRight) ? "5" : model.KneeExtensionRight)
                                          .Replace("#KneeFlexionLeft", string.IsNullOrEmpty(model.KneeFlexionLeft) ? "5" : model.KneeFlexionLeft)
                                          .Replace("#KneeFlexionRight", string.IsNullOrEmpty(model.KneeFlexionRight) ? "5" : model.KneeFlexionRight)
                                          .Replace("#AnkleDorsiLeft", string.IsNullOrEmpty(model.AnkleDorsiLeft) ? "5" : model.AnkleDorsiLeft)
                                          .Replace("#AnkleDorsiRight", string.IsNullOrEmpty(model.AnkleDorsiRight) ? "5" : model.AnkleDorsiRight)
                                          .Replace("#AnklePlantarLeft", string.IsNullOrEmpty(model.AnklePlantarLeft) ? "5" : model.AnklePlantarLeft)
                                          .Replace("#AnklePlantarRight", string.IsNullOrEmpty(model.AnklePlantarRight) ? "5" : model.AnklePlantarRight)
                                          .Replace("#EHLLeft", string.IsNullOrEmpty(model.EHLLeft) ? "5" : model.EHLLeft)
                                          .Replace("#EHLRight", string.IsNullOrEmpty(model.EHLRight) ? "5" : model.EHLRight);
            return htmlString;

        }

        public string GetROM(AIIntakeFormModel model, string bodyPart, string fileName)
        {
            string rom = "";

            var templateDir = Path.Combine(Environment.WebRootPath, "templates");

            //var filePath = templateDir + "/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId) + "/BHF";
            var filePath = templateDir + "/BHF/" + fileName;


            // 2. Read the entire file content into a string asynchronously
            string htmlString = System.IO.File.ReadAllText(filePath);


            switch (bodyPart.ToLower())
            {
                case "neck":
                    htmlString = htmlString.Replace("#ForwardFlexionROM", model.NeckFFROM)
                                           .Replace("#LRotationROM", model.NeckFFLeft)
                                           .Replace("#RRotationROM", model.NeckFFRight)
                                           .Replace("#ExtensionROM", model.NeckExt)
                                           .Replace("#LLateralFlexionROM", model.NeckExtLeft)
                                           .Replace("#RLateralFlexionROM", model.NeckExtRight)
                                           .Replace("#SPTest", model.NeckSpurlingTest);
                    break;
                case "midback":
                    htmlString = htmlString.Replace("#MbFFROM", model.MbFFROM)
                                           .Replace("#MbFFLeft", model.MbFFLeft)
                                           .Replace("#MbFFRight", model.MbFFRight)
                                           .Replace("#MbExt", model.MbExt);
                    break;
                case "lowback":
                    htmlString = htmlString.Replace("#LbFFROM", model.LBFFROM)
                                           .Replace("#LbFFLeft", model.LBFFLeft)
                                           .Replace("#LbFFRight", model.LBFFRight)
                                           .Replace("#LbExt", model.LBExt)
                                           .Replace("#SLRTest", model.LBSlrTest);
                    break;
                case "left shoulder":
                    htmlString = htmlString.Replace("#AbductionROM", model.LSHAbductionROM)
                                           .Replace("#ExternalRotationROM", model.LSHERROM)
                                           .Replace("#FlexionROM", model.LSHFlexionROM)
                                           .Replace("#InternalRotationROM", model.LSHIRROM)
                                           .Replace("#AdductionROM", model.LSHAdductionROM)
                                           .Replace("#ExtensionROM", model.LSHExtOM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right shoulder":
                    htmlString = htmlString.Replace("#AbductionROM", model.RSHAbductionROM)
                                           .Replace("#ExternalRotationROM", model.RSHERROM)
                                           .Replace("#FlexionROM", model.RSHFlexionROM)
                                           .Replace("#InternalRotationROM", model.RSHIRROM)
                                           .Replace("#AdductionROM", model.RSHAdductionROM)
                                           .Replace("#ExtensionROM", model.RSHExtOM)
                                           .Replace("#side", "RIGHT");
                    break;
                case "left knee":
                    htmlString = htmlString.Replace("#FlexionROM", model.LKNFlexionROM)
                                           .Replace("#ExtensionROM", model.LKNExtROM)
                                           .Replace("#FlexionNormal", model.LKNFlexionNormal)
                                           .Replace("#ExtensionNormal", model.LKNExtRNormal)
                                           .Replace("#FlexionTitle", model.LKNTitleROM)
                                           .Replace("#ExtensionTitle", model.LKNExtTitleROM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right knee":
                    htmlString = htmlString.Replace("#FlexionROM", model.RKNFlexionROM)
                                          .Replace("#ExtensionROM", model.RKNExtROM)
                                          .Replace("#FlexionNormal", model.RKNFlexionNormal)
                                          .Replace("#ExtensionNormal", model.RKNExtRNormal)
                                          .Replace("#FlexionTitle", model.RKNTitleROM)
                                          .Replace("#ExtensionTitle", model.RKNExtTitleROM)
                                          .Replace("#side", "RIGHT");
                    break;
                case "left ankle":
                    htmlString = htmlString.Replace("#PlantarFlexionROM", model.LAnkleFlexionROM)
                                           .Replace("#InversionROM", model.LAnkleInversionROM)
                                           .Replace("#DorsiflexionROM", model.LAnkleDorsiflexionROM)
                                           .Replace("#EversionROM", model.LAnkleEversionROM)
                                           .Replace("#side", "LEFT");

                    break;
                case "right ankle":
                    htmlString = htmlString.Replace("#PlantarFlexionROM", model.RAnkleFlexionROM)
                                           .Replace("#InversionROM", model.RAnkleInversionROM)
                                           .Replace("#DorsiflexionROM", model.RAnkleDorsiflexionROM)
                                           .Replace("#EversionROM", model.RAnkleEversionROM)
                                           .Replace("#side", "RIGHT");
                    break;
                case "left elbow":
                    htmlString = htmlString.Replace("#FlexionROM", model.LElbowFlexionROM)
                                           .Replace("#ExtensionROM", model.LElbowExtensionROM)
                                           .Replace("#SupinationROM", model.LElbowSupinationROM)
                                           .Replace("#PronationROM", model.LElbowPronationROM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right elbow":
                    htmlString = htmlString.Replace("#FlexionROM", model.RElbowFlexionROM)
                                           .Replace("#ExtensionROM", model.RElbowExtensionROM)
                                           .Replace("#SupinationROM", model.RElbowSupinationROM)
                                           .Replace("#PronationROM", model.RElbowPronationROM)
                                           .Replace("#side", "RIGHT");
                    break;
                case "left hip":
                    htmlString = htmlString.Replace("#FlexionROM", model.LHipFlexionROM)
                                           .Replace("#AdductionROM", model.LHipExtensionROM)
                                           .Replace("#ExtensionROM", model.LHipExtensionROM)
                                           .Replace("#InternalRotationROM", model.LHipIRotationROM)
                                           .Replace("#AbductionROM", model.LHipAbductionROM)
                                           .Replace("#ExternalRotationROM", model.LHipERotationROM)
                                           .Replace("#side", "LEFT");
                    break;
                case "right hip":
                    htmlString = htmlString.Replace("#FlexionROM", model.RHipFlexionROM)
                                           .Replace("#AdductionROM", model.RHipExtensionROM)
                                           .Replace("#ExtensionROM", model.RHipExtensionROM)
                                           .Replace("#InternalRotationROM", model.RHipIRotationROM)
                                           .Replace("#AbductionROM", model.RHipAbductionROM)
                                           .Replace("#ExternalRotationROM", model.RHipERotationROM)
                                            .Replace("#side", "RIGHT");
                    break;
                case "left wrist":
                    htmlString = htmlString.Replace("#FlexionROM", model.RWristFlexionROM)
                                           .Replace("#UlnarDeviationROM", model.RWristUDROM)
                                           .Replace("#ExtensionROM", model.RWristExtensionROM)
                                           .Replace("#RadialDeviationROM", model.RWristRDROM)
                                           .Replace("#side", "LEFT");

                    break;
                case "right wrist":
                    htmlString = htmlString.Replace("#FlexionROM", model.LWristFlexionROM)
                                           .Replace("#UlnarDeviationROM", model.LWristUDROM)
                                           .Replace("#ExtensionROM", model.LWristExtensionROM)
                                           .Replace("#RadialDeviationROM", model.LWristRDROM)
                                             .Replace("#side", "RIGHT");
                    break;
                default:
                    rom = "No ROM data available for this body part.";
                    break;
            }
            return htmlString;
        }
        #endregion


    }
}
