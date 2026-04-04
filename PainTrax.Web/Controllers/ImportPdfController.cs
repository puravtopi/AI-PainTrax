using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Humanizer;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Models;
using MS.Services;
using Newtonsoft.Json;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace PainTrax.Web.Controllers
{
    public class ImportPdfController : Controller
    {
        private readonly IWebHostEnvironment _environment;
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly PatientService _patientservices = new PatientService();
        private readonly ParentService _pareentservices = new ParentService();
        private readonly Common _commonservices = new Common();
        private readonly InscosService _inscosservices = new InscosService();
        private readonly AttorneysService _attorneyservices = new AttorneysService();
        private readonly AadjusterService _aadjusterService = new AadjusterService();
        private readonly EmpService _empService = new EmpService();
        private readonly UserService _userService = new UserService();

        public ImportPdfController(IWebHostEnvironment environment)
        {
            _environment = environment;

        }

        private int calculateAge(DateTime bday)
        {
            DateTime today = DateTime.Today;

            int age = today.Year - bday.Year;

            if (today.Month < bday.Month ||
       ((today.Month == bday.Month) && (today.Day < bday.Day)))
            {
                age--;  //birthday in current year not yet reached, we are 1 year younger ;)
                        //+ no birthday for 29.2. guys ... sorry, just wrong date for birth
            }

            return age;
        }


        #region njosmi

        protected string[] GetPdfData(string filename)
        {
            StringBuilder text = new StringBuilder();

            if (System.IO.File.Exists(filename))
            {
                using (PdfReader reader = new PdfReader(filename))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                    {
                        var strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(
                             pdfDoc.GetPage(page),
                            strategy
                         );
                        //var finder = new TextMarginFinder();
                        //string currentText=PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), finder);
                        currentText = System.Text.RegularExpressions.Regex.Replace(currentText, @" {4,}", "\n");
                        text.Append(currentText);
                    }
                }
            }

            string cont = text.ToString();
            string[] lines = cont.Split(new[] { '\n' });
            return lines;
        }

        public IActionResult Index()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
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

            //ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;
            return View();
        }

        [HttpPost]
        public IActionResult Index(List<IFormFile> files, string provider, string locationid)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;
            string message = "";
            if (locationid == null || locationid == "0")
            {
                message += $"<span class='text-danger'>Location not selected <span><br>";
                TempData["Message"] = message;
                return View();
            }

            var importFolder = Path.Combine(_environment.WebRootPath);
            if (files != null)
            {

                string data = "";
                foreach (var file in files)
                {
                    if (file.Length > 0 && Path.GetExtension(file.FileName).ToLower() == ".pdf")
                    {
                        string filePath = Path.Combine(importFolder, "temp" + Path.GetExtension(file.FileName).ToLower());

                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            file.CopyTo(stream);
                        }
                        string[] lines = GetPdfData(filePath);
                        var pdfdata = ExtractKeyValuePairs(lines);
                        TempData["pdfdata"] = JsonConvert.SerializeObject(pdfdata);
                        message += $"<span class='text-primary'>{file.FileName} File Processed Successfully<span><br>";
                        var result = SaveDetails(pdfdata, provider, locationid);
                        if (result == 1)
                            message += $"<span class='text-primary'>{file.FileName} File Import Successfully<span><br>";
                        else if (result == 2)
                            message += $"<span class='text-danger'>{file.FileName} File data already imported <span><br>";
                        else
                            message += $"<span class='text-danger'>{file.FileName} File not imported <span><br>";
                        data += string.Join("<br> ", lines);

                    }
                    else
                    {
                        message += $"<span class='text-danger'>{file.FileName} File blank or not docx<span><br>";
                    }
                }
                TempData["Message"] = message.ToString();
                TempData["data"] = data.ToString();
                TempData["alert"] = "alert alert-success";

            }
            else
            {
                TempData["Message"] = "File Not Uploaded.";
            }
            return View();

        }


        public int SaveDetails(List<KeyValuePair<string, string>> data, string provider, string locationid)
        {
            try
            {
                int patientId = 0, priminsId = 0, secinsId = 0, attornyId = 0, adjusterId = 0, empId = 0;

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                int age = calculateAge(DateTime.ParseExact(data.FirstOrDefault(x => x.Key == "D.O.B").Value, "MM/dd/yyyy", null));
                var nameValue = data.FirstOrDefault(x => x.Key == "Name").Value;

                tbl_patient objPatient = new tbl_patient()
                {
                    account_no = data.FirstOrDefault(x => x.Key == "MRN").Value,
                    address = data.FirstOrDefault(x => x.Key == "Address").Value,
                    dob = DateTime.ParseExact(data.FirstOrDefault(x => x.Key == "D.O.B").Value, "MM/dd/yyyy", null),
                    email = data.FirstOrDefault(x => x.Key == "Email").Value,
                    fname = nameValue.Split(' ')[0],
                    gender = data.FirstOrDefault(x => x.Key == "Sex").Value == "Male" ? "1" : data.FirstOrDefault(x => x.Key == "Sex").Value == "Female" ? "2" : "3",
                    home_ph = data.FirstOrDefault(x => x.Key == "HomePhone").Value,
                    lname = nameValue.Split(' ').Length > 1 ? nameValue.Split(' ')[^1] : "",
                    mobile = data.FirstOrDefault(x => x.Key == "MobilePhone").Value,
                    ssn = data.FirstOrDefault(x => x.Key == "SSN").Value,
                    cmp_id = cmpid,
                    createdby = userid,

                    age = age
                };
                List<tbl_patient> ptdata = new List<tbl_patient>();
                var dobStr = objPatient.dob.HasValue ? objPatient.dob.Value.ToString("yyyy-MM-dd") : "";
                var query = " and account_no='" + objPatient.account_no + "' and fname='" + objPatient.fname + "' and lname='" + objPatient.lname + "' and dob='" + dobStr + "' and cmp_id=" + cmpid + "";
                ptdata = _patientservices.GetAll(query);
                if (ptdata.Count > 0)
                {
                    return 2;
                }
                else
                {
                    patientId = _patientservices.Insert(objPatient);
                }
                HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, patientId);
                string pid = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
                ViewBag.patientId = patientId;
                query = "";
                List<tbl_inscos> insdata = new List<tbl_inscos>();
                tbl_inscos objInscos = new tbl_inscos();
                var addressList = data
                    .Where(x => x.Key.Equals("Address", StringComparison.OrdinalIgnoreCase))
                    .Select(x => x.Value)
                    .ToList();
                string carrier = data.FirstOrDefault(x => x.Key == "Carrier").Value;
                string address = addressList.Count > 1 ? addressList[1] : null;

                if (!string.IsNullOrEmpty(carrier))
                {
                    query = " and cmpname='" + carrier + "' and cmp_id=" + cmpid + "";
                    insdata = _inscosservices.GetAll(query);

                    objInscos = new tbl_inscos()
                    {
                        address1 = address,
                        cmpname = carrier,
                        telephone = data.FirstOrDefault(x => x.Key == "PhoneNumber").Value,
                        cmp_id = cmpid,
                        createdby = userid
                    };

                    if (insdata.Count > 0)
                    {
                        objInscos.id = insdata[0].id.Value;
                        _inscosservices.Update(objInscos);
                        priminsId = insdata[0].id.Value;
                    }
                    else
                    {
                        priminsId = _inscosservices.Insert(objInscos);
                    }
                }

                //save IE details 

                tbl_patient_ie objIE = new tbl_patient_ie()
                {
                    adjuster_id = adjusterId,
                    attorney_id = attornyId,
                    created_by = userid,

                    doe = DateTime.Now,
                    emp_id = empId,
                    is_active = true,
                    location_id = Convert.ToInt32(locationid),
                    patient_id = patientId,
                    primary_ins_cmp_id = priminsId,
                    primary_policy_no = data.FirstOrDefault(x => x.Key == "PolicyNo").Value,


                };
                int ie = 0;
                ie = _ieService.Insert(objIE);

                HttpContext.Session.SetInt32(SessionKeys.SessionIEId, ie);
                return 1;
                //                return Json(new { status = 1, patintid = patientId, ieid = ie });

            }
            catch (Exception ex)
            {
                return 0;
                //              return Json(new { status = 0 });
            }


        }




        public static List<KeyValuePair<string, string>> ExtractKeyValuePairs(string[] lines)
        {
            var pairs = new List<KeyValuePair<string, string>>();

            // This regex finds *all* key:value pairs per line
            string pattern = @"([A-Za-z0-9\/\.\(\)\- ]+?):\s*([^:]*)(?=\s+[A-Za-z0-9\/\.\(\)\- ]+?:|$)";
            foreach (string pdfText in lines)
            {
                string text = pdfText
                .Replace("Home Phone", "HomePhone")
                .Replace("Work Phone", "WorkPhone")
                .Replace("Mobile Phone", "MobilePhone")
                .Replace("Preferred Contact Method", "PreferredContactMethod")
                .Replace("PMS ID", "PMSID")
                .Replace("Marital Status", "MaritalStatus")
                .Replace("Emergency Contact", "EmergencyContact")
                .Replace("Emergency Contact Phone", "EmergencyContactPhone")
                .Replace("Primary Care Provider", "PrimaryCareProvider")
                .Replace("Referring Providers", "ReferringProviders")
                .Replace("Policy #", "PolicyNo")
                .Replace("Group ID/Name", "GroupID/Name")
                .Replace("Phone Number", "PhoneNumber")
                .Replace("Fax Number", "FaxNumber");

                foreach (var rawLine in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    foreach (Match match in Regex.Matches(line, pattern))
                    {
                        string key = match.Groups[1].Value.Trim();
                        string value = match.Groups[2].Value.Trim();

                        pairs.Add(new KeyValuePair<string, string>(key, value));
                    }
                }
            }
            return pairs;

        }


        #endregion

        #region QMPCC
        [HttpPost]
        public IActionResult QMPPC(List<IFormFile> files, string provider, string locationid)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;
            string message = "";
            if (locationid == null || locationid == "0")
            {
                message += $"<span class='text-danger'>Location not selected <span><br>";
                TempData["Message"] = message;
                return RedirectToAction("Index");
            }

            var importFolder = Path.Combine(_environment.WebRootPath);
            if (files != null)
            {

                string data = "";
                foreach (var file in files)
                {
                    if (file.Length > 0 && Path.GetExtension(file.FileName).ToLower() == ".pdf")
                    {
                        string filePath = Path.Combine(importFolder, "temp" + Path.GetExtension(file.FileName).ToLower());

                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            file.CopyTo(stream);
                        }
                        string[] lines = GetPdfDataQMPPC(filePath);
                        var pdfdata = ExtractKeyValuePairsQMPPC(lines);
                        TempData["pdfdata"] = JsonConvert.SerializeObject(pdfdata);
                        message += $"<span class='text-primary'>{file.FileName} File Processed Successfully<span><br>";
                        var result = SaveDetailsQMPPC(pdfdata, provider, locationid);
                        if (result == 1)
                            message += $"<span class='text-primary'>{file.FileName} File Import Successfully<span><br>";
                        else if (result == 2)
                            message += $"<span class='text-danger'>{file.FileName} File data already imported <span><br>";
                        else
                            message += $"<span class='text-danger'>{file.FileName} File not imported <span><br>";
                        data += string.Join("<br> ", lines);

                    }
                    else
                    {
                        message += $"<span class='text-danger'>{file.FileName} File blank or not docx<span><br>";
                    }
                }
                TempData["Message"] = message.ToString();
                TempData["data"] = data.ToString();
                TempData["alert"] = "alert alert-success";

            }
            else
            {
                TempData["Message"] = "File Not Uploaded.";
            }
            return RedirectToAction("Index");

        }


        public static List<KeyValuePair<string, string>> ExtractKeyValuePairsQMPPC(string[] lines)
        {
            var pairs = new List<KeyValuePair<string, string>>();

            string[] knownKeys =
            {
        "First Name","Middle Name","Last Name","D.O.B","Gender","Address",
        "City","State","Zip",
        "Cell Phone #","Home Phone","Work","Email","Extn.",
        "Attorney Address","Attorney Phone","Attorney Email","Attorney Fax",
        "Attorney","Case Type",
        "Case Status","SSN",
        "Policy Holder","Name","Phone","Fax",
        "Contact Person","Claim File #","Policy #","WCB #",
        "Accident Date","Plate Number","Report Number",
        "Hospital Name","Hospital Address","Date of Admission",
        "Additional Patient","Describe Injury","Patient Type",
        "Date of First Treatment","Chart #","Extension"
    };

            // Sort longest first (important for Attorney vs Attorney Address)
            var sortedKeys = knownKeys
                .OrderByDescending(k => k.Length)
                .Select(Regex.Escape);

            string keyPattern = string.Join("|", sortedKeys);

            // 🔥 FIXED REGEX (no word boundary)
            string pattern = $@"({keyPattern})\s+(.*?)(?=\s+({keyPattern})\s+|$)";

            foreach (string pdfText in lines)
            {
                if (string.IsNullOrWhiteSpace(pdfText))
                    continue;

                foreach (var rawLine in pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    foreach (Match match in Regex.Matches(line, pattern))
                    {
                        string key = match.Groups[1].Value.Trim();
                        string value = match.Groups[2].Value.Trim();

                        pairs.Add(new KeyValuePair<string, string>(key, value));
                    }
                }
            }

            return pairs;
        }
        protected string[] GetPdfDataQMPPC(string filename)
        {
            StringBuilder text = new StringBuilder();

            if (System.IO.File.Exists(filename))
            {
                using (PdfReader reader = new PdfReader(filename))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                    {
                        var strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(
                             pdfDoc.GetPage(page),
                            strategy
                         );
                        //var finder = new TextMarginFinder();
                        //string currentText=PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), finder);
                        //  currentText = System.Text.RegularExpressions.Regex.Replace(currentText, @" {4,}", "\n");
                        text.Append(currentText);
                    }
                }
            }

            string cont = text.ToString();
            string[] lines = cont.Split(new[] { '\n' });
            return lines;
        }

        public int SaveDetailsQMPPC(List<KeyValuePair<string, string>> data, string provider, string locationid)
        {
            try
            {
                int patientId = 0, priminsId = 0, secinsId = 0, attornyId = 0, adjusterId = 0, empId = 0;

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                int age = calculateAge(DateTime.ParseExact(data.FirstOrDefault(x => x.Key == "D.O.B").Value, "MM/dd/yyyy", null));
                var nameValue = data.FirstOrDefault(x => x.Key == "Name").Value;
                string pstate = "";
                string statename = data.FirstOrDefault(x => x.Key == "State").Value.Trim('-');
                System.Data.DataTable dt = _pareentservices.GetData($"select state_name,fullname from tbl_state where fullname = '{statename}' ");
                if (dt.Rows.Count > 0)
                {

                    try { pstate = dt.Rows[0]["state_name"].ToString(); } catch { }

                }



                tbl_patient objPatient = new tbl_patient()
                {

                    fname = data.FirstOrDefault(x => x.Key == "First Name").Value,
                    mname = data.FirstOrDefault(x => x.Key == "Middle Name").Value.Trim('-'),
                    lname = data.FirstOrDefault(x => x.Key == "Last Name").Value.Trim('-'),
                    dob = DateTime.ParseExact(data.FirstOrDefault(x => x.Key == "D.O.B").Value, "MM/dd/yyyy", null),
                    address = data.FirstOrDefault(x => x.Key == "Address").Value.Trim('-'),
                    city = data.FirstOrDefault(x => x.Key == "City").Value.Trim('-'),
                    state = pstate,
                    zip = data.FirstOrDefault(x => x.Key == "Zip").Value.Trim('-'),
                    email = data.FirstOrDefault(x => x.Key == "Email").Value.Trim('-'),
                    mobile = data.FirstOrDefault(x => x.Key == "Cell Phone #").Value.Trim('-'),
                    home_ph = data.FirstOrDefault(x => x.Key == "Home Phone").Value.Trim('-'),
                    gender = data.FirstOrDefault(x => x.Key == "Gender").Value == "Male" ? "1" : data.FirstOrDefault(x => x.Key == "Gender").Value == "Female" ? "2" : "3",
                    ssn = data.FirstOrDefault(x => x.Key == "SSN").Value,
                    cmp_id = cmpid,
                    createdby = userid,
                    mc = "No",
                    vaccinated = false,
                    age = age

                };
                List<tbl_patient> ptdata = new List<tbl_patient>();
                var dobStr = objPatient.dob.HasValue ? objPatient.dob.Value.ToString("yyyy-MM-dd") : "";
                var query = "and fname='" + objPatient.fname + "' and lname='" + objPatient.lname + "' and dob='" + dobStr + "' and cmp_id=" + cmpid + "";
                ptdata = _patientservices.GetAll(query);
                if (ptdata.Count > 0)
                {
                    return 2;
                }
                else
                {
                    patientId = _patientservices.Insert(objPatient);
                }
                HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, patientId);
                string pid = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
                ViewBag.patientId = patientId;
                query = "";
                List<tbl_inscos> insdata = new List<tbl_inscos>();
                List<tbl_attorneys> attdata = new List<tbl_attorneys>();

                tbl_inscos objInscos = new tbl_inscos();
                tbl_attorneys objAttorney = new tbl_attorneys();
                var addressList = data
                    .Where(x => x.Key.Equals("Address", StringComparison.OrdinalIgnoreCase))
                    .Select(x => x.Value)
                    .ToList();

                var cityList = data
                    .Where(x => x.Key.Equals("City", StringComparison.OrdinalIgnoreCase))
                    .Select(x => x.Value)
                    .ToList();

                var stateList = data
                    .Where(x => x.Key.Equals("State", StringComparison.OrdinalIgnoreCase))
                    .Select(x => x.Value)
                    .ToList();

                var zipList = data
                    .Where(x => x.Key.Equals("Zip", StringComparison.OrdinalIgnoreCase))
                    .Select(x => x.Value)
                    .ToList();

                string insco = data.FirstOrDefault(x => x.Key == "Name").Value;
                string attorney = data.FirstOrDefault(x => x.Key == "Attorney").Value;

                if (!string.IsNullOrEmpty(insco.Trim('-')))
                {
                    query = " and LOWER(cmpname)=LOWER('" + insco + "') and cmp_id=" + cmpid + "";
                    insdata = _inscosservices.GetAll(query);
                    string insaddress = string.Join(", ", new[]
                           {
                                addressList.Count > 1 ? addressList[1].Trim('-') : null,
                                cityList.Count > 1 ? cityList[1].Trim('-') : null,
                                stateList.Count > 1 ? stateList[1].Trim('-') : null,
                                zipList.Count > 1 ? zipList[1].Trim('-') : null
                            }.Where(x => !string.IsNullOrWhiteSpace(x)));

                    objInscos = new tbl_inscos()
                    {

                        address1 = insaddress,
                        city = cityList.Count > 1 ? cityList[1].Trim('-') : null,
                        state = stateList.Count > 1 ? stateList[1].Trim('-') : null,
                        // zip = zipList.Count > 1 ? zipList[1].Trim('-') : null,

                        cmpname = data.FirstOrDefault(x => x.Key == "Name").Value.Trim('-'),
                        telephone = data.FirstOrDefault(x => x.Key == "Phone").Value.Trim('-'),
                        faxno = data.FirstOrDefault(x => x.Key == "Fax").Value.Trim('-'),
                        cmp_id = cmpid,
                        createdby = userid
                    };

                    if (insdata.Count > 0)
                    {
                        objInscos.id = insdata[0].id.Value;
                        // _inscosservices.Update(objInscos);
                        priminsId = insdata[0].id.Value;
                    }
                    else
                    {
                        priminsId = _inscosservices.Insert(objInscos);
                    }
                }

                if (!string.IsNullOrEmpty(attorney.Trim('-')))
                {
                    query = " and LOWER(attorney)=LOWER('" + attorney + "') and cmp_id=" + cmpid + "";
                    attdata = _attorneyservices.GetAll(query);

                    objAttorney = new tbl_attorneys()
                    {
                        Attorney = attorney,
                        Address = data.FirstOrDefault(x => x.Key == "Attorney Address").Value.Trim('-'),
                        EmailId = data.FirstOrDefault(x => x.Key == "Attorney Email").Value.Trim('-'),
                        ContactNo = data.FirstOrDefault(x => x.Key == "Attorney Phone").Value.Trim('-'),
                        cmp_id = cmpid,
                        CreatedBy = userid
                    };

                    if (attdata.Count > 0)
                    {
                        objAttorney.Id = attdata[0].Id.Value;
                        //_attorneyservices.Update(objAttorney);

                        attornyId = attdata[0].Id.Value;
                    }
                    else
                    {
                        attornyId = _attorneyservices.Insert(objAttorney);
                    }
                }


                //save IE details 

                tbl_patient_ie objIE = new tbl_patient_ie()
                {
                    adjuster_id = adjusterId,
                    attorney_id = attornyId,
                    created_by = userid,
                    compensation = data.FirstOrDefault(x => x.Key == "Case Type").Value == "NoFault" ? "NF" : "",
                    doa = DateTime.ParseExact(data.FirstOrDefault(x => x.Key == "Accident Date").Value, "MM/dd/yyyy", null),
                    doe = DateTime.Now,
                    emp_id = empId,
                    is_active = true,
                    location_id = Convert.ToInt32(locationid),
                    patient_id = patientId,
                    primary_ins_cmp_id = priminsId,
                    primary_policy_no = data.FirstOrDefault(x => x.Key == "Policy #").Value.Trim('-'),
                    primary_claim_no = data.FirstOrDefault(x => x.Key == "Claim File #").Value.Trim('-'),


                };
                int ie = 0;
                ie = _ieService.Insert(objIE);

                HttpContext.Session.SetInt32(SessionKeys.SessionIEId, ie);
                return 1;
                //                return Json(new { status = 1, patintid = patientId, ieid = ie });

            }
            catch (Exception ex)
            {
                return 0;
                //              return Json(new { status = 0 });
            }


        }



        #endregion

    }

}
