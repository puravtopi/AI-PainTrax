using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using MS.Models;
using MS.Services;
using System.Data;
using PainTrax.Services;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class AppointmentController : Controller
    {
        private readonly Common _commonservices = new Common();
        private readonly AppHelper _apphelper = new AppHelper();
        private readonly ParentService _parentService = new ParentService();
        private readonly PatientService _patientservices = new PatientService();
        private readonly AppointmentService _appointmentservice = new AppointmentService();
        private readonly AppStatusService _appStatusService = new AppStatusService();
        private readonly AppProviderService _appProviderService = new AppProviderService();
        private readonly AppProviderRelService _appProviderRelService = new AppProviderRelService();




        public IActionResult Index()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View();
        }


        public IActionResult Appointments(int id = 0, string searchStart = "", string searchEnd = "")
        {
            var patientData = _patientservices.GetOne(id);
            DataTable dt = new DataTable();
            if (searchStart != "")
                dt = _patientservices.GetData("select * from view_reportappointment where patient_id=" + id + " and appointmentDate>='" + searchStart + "' and appointmentDate<='" + searchEnd + "' order by appointmentDate");
            else
                dt = _patientservices.GetData("select * from view_reportappointment where patient_id=" + id + " order by appointmentDate");
            ViewBag.list = dt.AsEnumerable();
            return View(patientData);
        }

        public IActionResult New(int patientid = 0, int appid = 0, string type = "",bool caladd=false, string selecteddate="",string selectedtime="",int locationid=0)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            ViewBag.providerList = _appProviderService.GetAllCompany(cmpid.Value); 
            var data = new tbl_appointment();
            data.location_id = locationid;
            ViewBag.btnTitle = "Set";
            ViewBag.btnClass = "btn-success";
            // ViewBag.patientList = _patientservices.GetPatientList(cmpid.Value);
            //data.appointmentStart = selectedtime.TrimStart('0');
            if (caladd)
            {
                ViewBag.readOnly = "";
                
                if (type == "multiple")
                {
                    ViewBag.btnTitle = "Set Multiple";
                    ViewBag.type = "multiple";
                }
                else
                {
                    data.appointmentDate = selecteddate.Substring(0, 4) + "-" + selecteddate.Substring(4, 2) + "-" + selecteddate.Substring(6, 2);
                    data.appointmentStart= DateTime.Parse( selectedtime).ToString("HH:mm");
                }
            }
            else 
            { 
                if (patientid > 0)
                {
                    var patientData = _patientservices.GetOne(patientid);
                    ViewBag.patientData = patientData;
                    ViewBag.readOnly = "readonly";
                    data.appointmentDate = DateTime.Now.ToString("yyyy-MM-dd");
                    if (type == "multiple")
                    {
                        ViewBag.btnTitle = "Set Multiple";
                        ViewBag.type = "multiple";
                    }
                }
                else if (appid > 0)
                {
                    string []provider = GetAppProviders(appid);
                    data = _appointmentservice.GetOne(appid);
                    var patientData = _patientservices.GetOne(Convert.ToInt32(data.patient_id));
                    ViewBag.patientData = patientData;
                    ViewBag.readOnly = "readonly";
                    ViewBag.providerids = provider[0];
                    ViewBag.providernames = provider[1];
                    ViewBag.type = type;
                    if (type == "update")
                    {
                        ViewBag.btnTitle = "Update";
                        ViewBag.btnClass = "btn-primary";
                    }
                    else if (type == "remove")
                    {
                        ViewBag.btnTitle = "Delete";
                        ViewBag.btnClass = "btn-danger";
                    }

                }
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult SaveAppointments(AppointmentVM data)
        {
            //var newData = new tbl_appointment();
            //newData.status_id = 1;
            //newData.patient_id = data.patient_id;
            //newData.patientIE_id = data.patientIE_id;
            //newData.location_id = data.location_id;
            //newData.appointmentStart = data.appointmentStart;
            //newData.appointmentDate = data.appointmentDate;
            //newData.appointmentNote = data.appointmentNote;
            //newData.appointmentEnd = DateTime.ParseExact(data.appointmentStart, "HH:mm", null).AddMinutes(30).ToString("HH:mm");
            // _appointmentservice.Insert(newData);

            data.status_id = 1;
            data.appointmentEnd = DateTime.ParseExact(data.appointmentStart, "HH:mm", null).AddMinutes(30).ToString("HH:mm");
            _appointmentservice.InsertWithProvider(data);
            return Json(new { status = 1 });
        }

        [HttpPost]
        public IActionResult UpdateAppointments(AppointmentVM data)
        {
            //var updateData = _appointmentservice.GetOne(Convert.ToInt32(data.appointment_id));
            //updateData.appointmentDate = data.appointmentDate;
           // updateData.appointmentStart = data.appointmentStart;
          //  updateData.appointmentNote = data.appointmentNote;
            
            data.appointmentEnd = DateTime.ParseExact(data.appointmentStart, "HH:mm", null).AddMinutes(30).ToString("HH:mm");
            //_appointmentservice.Update(updateData);
            _appointmentservice.UpdateWithProvider(data);
            return Json(new { status = 1 });
        }

        [HttpPost]
        public IActionResult RemoveAppointments(AppointmentVM data)
        {
            var deleteData = new tbl_appointment();
            deleteData.appointment_id = data.appointment_id;
            _appointmentservice.Delete(deleteData);
            return Json(new { status = 1 });
        }

        [HttpPost]
        public IActionResult MultipleAppointments(MultiAppointment data)
        {
//            var dates = new List<DateTime>();
            string[] days = data.days.Split(new char[] { ',' });
            var dt = new DateTime();
            for (dt = Convert.ToDateTime(data.appointmentFromDate); dt <= Convert.ToDateTime(data.appointmentToDate); dt = dt.AddDays(1))
            {
                if (days.Contains(Convert.ToString(dt.DayOfWeek)))
                {
                    //var newData = new tbl_appointment();
                    var newData = new AppointmentVM();
                    newData.status_id = 1;
                    newData.patient_id = data.patient_id;
                    newData.patientIE_id = data.patientIE_id;
                    newData.location_id = data.location_id;
                    newData.appointmentStart = data.appointmentStart;
                    newData.appointmentEnd = DateTime.ParseExact(newData.appointmentStart, "HH:mm", null).AddMinutes(30).ToString("HH:mm");
                    newData.appointmentDate = dt.ToString("yyyy-MM-dd"); ;
                    newData.appointmentNote = data.appointmentNote;
                    newData.providers = data.providers;
                    //_appointmentservice.Insert(newData);
                    _appointmentservice.InsertWithProvider(newData);


                    //                  dates.Add(dt);
                }
            }


            return Json(new { status = 1 });
        }


        public IActionResult Reports()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            ViewBag.statusList = _appStatusService.GetAllDropDown();
            List<SelectListItem> provider = _appProviderService.GetAllCompany(cmpid.Value);
            provider.Insert(0, new SelectListItem("All", "0"));
            ViewBag.providerList = provider;
            return View();
        }
        public IActionResult ChangeStatus(int appid = 0, int statusId = 0)
        {
            ViewBag.appointmentId = appid;
            ViewBag.statusList = _appStatusService.GetAllDropDown(false, statusId);
            return View();
        }

        [HttpPost]
        public IActionResult ChangeStatus(tbl_appointment data)
        {
            _appointmentservice.UpdateStatus(data);
            return View();
        }


        public string GetAppointmentCounts(int location_id = 0)
        {
            string data = _apphelper.GetJson(@"select `date`,`appointments` from view_countappointments where Location_id=" + location_id.ToString());
            return data;

        }
        public string GetDayAppointments(string selected_date = "", int location_id = 0)
        {
            string data = _apphelper.GetJson(@"select * from view_reportappointment where REPLACE(appointmentDate,'-','')=" + selected_date + " and Location_id=" + location_id.ToString());
            return data;

        }

        public string GetReportAppointments(string from_date = "", string to_date = "", int location_id = 0,int status_id=0,int provider_id=0)
        {
            // string data = _apphelper.GetJson(@"select  * from view_reportappointment ");
            string data = _apphelper.GetJson(filterData(from_date , to_date, location_id ,status_id , provider_id));
            return data;

        }
        public IActionResult GetExcelAppointments(string from_date = "", string to_date = "", int location_id = 0, int status_id = 0, int provider_id = 0)
        {
            // string data = _apphelper.GetJson(@"select  * from view_reportappointment ");
            DataTable dt = _parentService.GetData (filterData(from_date, to_date, location_id, status_id, provider_id));
            byte[] excelBytes = _apphelper.GetExcel(dt);
            return File(excelBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.spreedsheet");

        }

        public IActionResult GetPdfAppointments(string from_date = "", string to_date = "", int location_id = 0, int status_id = 0, int provider_id = 0)
        {
            // string data = _apphelper.GetJson(@"select  * from view_reportappointment ");
            DataTable dt = _parentService.GetData(filterData(from_date, to_date, location_id, status_id, provider_id));
            byte[] pdfBytes = _apphelper.GetPdf(dt);
            return File(pdfBytes, "application/pdf");

        }

        public JsonResult SearchPatients(string prefix)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            List<string> _patients = new List<string>();
            _patients = _patientservices.GetPatientSearchList(cmpid.Value, prefix);
            return Json(_patients);

        }
        public string[] GetAppProviders(int appointment_id)
        {
            string[] providers = new string[] { "",""};
            DataTable dt=_parentService.GetData("select * from view_appprovider where app_id="+appointment_id );
            if(dt.Rows.Count >0)
            {
                providers[0] = dt.Rows[0]["ProviderIds"].ToString();
                providers[1] = dt.Rows[0]["ProviderName"].ToString();
            }
            return providers; 
        }

        [HttpPost]
        public IActionResult TransferDate(TransferVM data)
        {
            _appointmentservice.Transfer(data);
            return Json(new { status = 1 });
        }

        protected string filterData(string from_date = "", string to_date = "", int location_id = 0, int status_id = 0, int provider_id = 0)
        {

            string ProviderIds = "";
        /*    foreach (System.Web.UI.WebControls.ListItem item in lstProvider.Items)
            {
                if (item.Selected)
                {
                    ProviderIds += item.Value + ",";
                }
            }
            ProviderIds = ProviderIds.Trim(new char[] { ',' });*/

            string cond = "";
            if (status_id> 0)
            {
                cond += " and Status_Id=" + status_id+ " ";
            }
            if (location_id> 0)
            {
                cond += " and Location_Id=" + location_id  + " ";
            }
            if (from_date !="" && from_date != null)
            {
                DateTime fdate = Convert.ToDateTime(from_date);
                cond += " and AppointmentDate>='" + fdate.ToString("yyyy-MM-dd") + "'";
            }
            if (to_date !="" && to_date != null)
            {
                DateTime tdate = Convert.ToDateTime(to_date);
                cond += " and AppointmentDate<='" + tdate.ToString("yyyy-MM-dd") + "'";
            }

            //if(txtFDate.Value.Length > 0 and txtT)
            if (cond.Length > 0)
            {
                cond = cond.Substring(4);
                cond = " where " + cond;
            }
            //if (ProviderIds.Length == 0)
            if (provider_id == 0)
                return "select * from view_reportappointment  " + cond + " order by AppointmentDate desc,AppointmentStart desc";
            else
                return "select * from view_reportappointment  " + (cond.Length == 0 ? " Where " : cond + " and ") + "Appointment_Id in (select App_Id from tbl_appProvider_rel where app_provider_id in (" + provider_id + ")) order by AppointmentDate desc,AppointmentStart desc";
        }



        //public void SaveProviders(int appid,string appproviders)
        //{
        //    string[] appproviderlist = appproviders.Split(new char[] { ',' });
        //    _appProviderRelService.DeleteByAppointmentId(appid);
        //    foreach( var provider in appproviderlist)
        //    {
        //        tbl_appprovider_rel data = new tbl_appprovider_rel();
        //        data.appid = appid;
        //        data.app_provider_id = Convert.ToInt16(provider);
        //        _appProviderRelService.Insert(data);
        //    }
        //}
    }
}
