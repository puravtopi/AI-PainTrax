using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using MS.Services;
using Optivem.Framework.Core.Domain;
using PainTrax.Web.AzureServices;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Text.RegularExpressions;

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
        public IActionResult Create(int? locId, int? id)
        {
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
            ViewBag.CmpId = cmpid.ToString();

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
                return PartialView("_IntakeBHF");
            else if (client_code.ToLower() == "hposm")
                return PartialView("_IntakeHPOSM");
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
    }
}
