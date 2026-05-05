using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Services;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class AppointmentsController : Controller
    {
        private readonly Common _commonservices = new Common();
        private readonly AppHelper _apphelper = new AppHelper();
        private readonly ParentService _parentService = new ParentService();
        private readonly PatientService _patientservices = new PatientService();
        private readonly AppointmentService _appointmentservice = new AppointmentService();
        private readonly AppStatusService _appStatusService = new AppStatusService();
        private readonly AppProviderService _appProviderService = new AppProviderService();
        private readonly AppProviderRelService _appProviderRelService = new AppProviderRelService();
        private readonly UserService _userService = new UserService();

        public IActionResult Index()
        {
       
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var locdata = _commonservices.GetLocations(Convert.ToInt32(cmpid));
            List<SelectListItem> lst = new List<SelectListItem>();
            int defaultlocation = HttpContext.Session.GetInt32(SessionKeys.SessionLocationId).Value;
            foreach (var item in locdata)
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
            providers.RemoveAt(0);
            ViewBag.provList = providers;
         
            return View();
        }

        public IActionResult ChangeStatus(int? appid, int statusid)
        {
            ViewBag.appid = appid;
            ViewBag.statusList = _appStatusService.GetAllDropDown(false, statusid);
            return PartialView("_ChangeStatus");
        }

        [HttpPost]
        public IActionResult Move(int? appid, string app_date , string app_time)
        {
            ViewBag.appid = appid;
            ViewBag.app_date = app_date;
            ViewBag.app_time = app_time;
            return PartialView("_Move");
        }

        public IActionResult Appointment(int? id, int providerId, string date, string time,int locationid, string mode="")
        {
            AppointmentsVM model;
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            
            var locdata = _commonservices.GetLocations(Convert.ToInt32(cmpid));
            List<SelectListItem> lst = new List<SelectListItem>();
            
            foreach (var item in locdata)
            {
                var obj = new SelectListItem()
                {
                    Text = item.Text,
                    Value = item.Value,
                    Selected = item.Value == locationid.ToString() ? true : false
                };
                lst.Add(obj);

            }
            ViewBag.locList = lst;
            if (id.HasValue)
            {
                
                model = _appointmentservice.GetOneNew(id ?? 0);
                var patientData = _patientservices.GetOne(Convert.ToInt32(model.patient_id));
                ViewBag.patientData = patientData;
                model.isEdit = 1;
            }
            else
            {
                // ADD MODE
                model = new AppointmentsVM
                {
                    provider_id = providerId,
                    location_id= locationid,
                    app_date = date,
                    app_time = time,
                    isEdit = 0
                };                
                
            }
            
            return PartialView("_Appointment", model);
        }

        public JsonResult SearchPatients(string prefix)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            List<string> _patients = new List<string>();
            _patients = _patientservices.GetPatientSearchList(cmpid.Value, prefix);
            return Json(_patients);

        }

        [HttpPost]
        public IActionResult AppointmentSave(AppointmentsVM data)
        {
            
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            data.cmp_id = cmpid;
            if (data.isEdit==0)
            {
                data.status_id = 1;                
                _appointmentservice.InsertNew(data);
               
            }
            else
            {
                _appointmentservice.UpdateNew(data);
            }
            return RedirectToAction("Index");

        }

        [HttpPost]
        public IActionResult DeleteAppointment(int appid)
        {
            AppointmentsVM data=new AppointmentsVM();   
            data.app_id = appid;
            _appointmentservice.DeleteNew(data); 
            return Json(new { success = true });
        }
        
        [HttpPost]
        public IActionResult ChangeStatusSave(int appid, int statusid)
        {
            AppointmentsVM data = new AppointmentsVM();
            data.app_id = appid;
            data.status_id = statusid;
            _appointmentservice.UpdateStatusNew(data);
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult MoveSave(int appid, string  appdate , string apptime )
        {
            AppointmentsVM data = new AppointmentsVM();
            data.app_id = appid;
            data.app_date = appdate;
            data.app_time = apptime;
            _appointmentservice.Move(data);
            return RedirectToAction("Index");
        }

        [HttpGet]
        public string GetAppointmentsDay(string selected_date, int cmp_id, int location_id, string provider_ids)
        {
            string query = @"
                            SELECT 
                                a.app_id AS id,
                                a.provider_id AS providerId,
                                a.location_id,
                                a.cmp_id,

                                CONCAT(p.fname, ' ', p.lname , ' | ' , a.app_note ) AS title,

                                STR_TO_DATE(CONCAT(a.app_date, ' ', a.app_time), '%Y-%m-%d %H:%i') AS start,

                                DATE_ADD(
                                    STR_TO_DATE(CONCAT(a.app_date, ' ', a.app_time), '%Y-%m-%d %H:%i'),
                                    INTERVAL 30 MINUTE
                                ) AS end,

                                s.status_id,
                                s.status,

                                0 AS count

                            FROM tbl_appointments a
                            LEFT JOIN tbl_patient p ON p.id = a.patient_id
                            LEFT JOIN tbl_app_status s ON s.status_id = a.status_id

                            WHERE a.cmp_id = " + cmp_id + @"
                            AND a.location_id = " + location_id + @"
                            AND REPLACE(a.app_date,'-','') = " + selected_date + @"
                            AND a.provider_id IN (" + provider_ids + ")";

            return _apphelper.GetJson(query);
        }

        [HttpGet]
        public string GetAppointmentsMonth(int cmp_id, int location_id, string provider_ids)
        {
            string query = @"
                            SELECT 
                                0 AS id,
                                cmp_id,
                                provider_id AS providerId,
                                location_id,

                                '' AS title,
                                '' AS status,

                                STR_TO_DATE(CONCAT(app_date, ' ', app_time), '%Y-%m-%d %H:%i') AS start,
                                STR_TO_DATE(CONCAT(app_date, ' ', app_time), '%Y-%m-%d %H:%i') AS end,

                                COUNT(*) AS count

                            FROM tbl_appointments
                            WHERE cmp_id = " + cmp_id + @"
                            AND location_id = " + location_id + @"
                            AND provider_id IN (" + provider_ids + @")

                            GROUP BY cmp_id, provider_id, location_id, app_date";
            return _apphelper.GetJson(query);
        }
    }
}
