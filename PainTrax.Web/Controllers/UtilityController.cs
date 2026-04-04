using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewComponents;
using MS.Models;
using MS.Services;
using Optivem.Framework.Core.Domain;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class UtilityController : Controller
    {

        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;
        private readonly SignatureService _services = new SignatureService();
        private readonly CompanyServices _cmpservices = new CompanyServices();
        private readonly PatientService _patientservices = new PatientService();
        private readonly ILogger<UtilityController> _logger;
        private readonly PatientImportService _patientimportservice = new PatientImportService();

        public UtilityController(Microsoft.AspNetCore.Hosting.IHostingEnvironment environment, ILogger<UtilityController> logger,
                               IConfiguration configuration)
        {

            Environment = environment;
            Configuration = configuration;
            _logger = logger;
        }

        public IActionResult UploadSign()
        {
            return View(new UploadSignVM());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UploadSign(IFormFile[] signs)
        {

            if (signs != null)
            {
                string folderPath = Path.Combine(Environment.WebRootPath, "signatures");
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                foreach (var item in signs)
                {
                    try
                    {

                        var fname = Path.GetFileNameWithoutExtension(item.FileName.Split(".")[0]);

                        // Generate a unique filename 
                        string fileName = fname.Split("_")[0] + Path.GetExtension(item.FileName);
                        string filePath = Path.Combine(folderPath, fileName);

                        using (var fileStream = new FileStream(filePath, FileMode.Create))
                        {
                            item.CopyTo(fileStream);
                        }

                        tbl_ie_sign model = new tbl_ie_sign()
                        {

                            patient_id = Convert.ToInt32(fname.Split("_")[0]),
                            signatureData = fileName
                        };
                        _services.ManageSign(model);
                    }
                    catch (Exception)
                    {

                    }
                    //model.header_template = fileName;
                }

            }
            TempData["msg"] = "<script>alert('Upload succesfully');</script>";
            return View();
        }

        public IActionResult UploadDocuments()
        {
            var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments");
            string[] dirs = Directory.GetDirectories(FolderPath, "*", SearchOption.TopDirectoryOnly);


            List<UploadDocVM> lst = new List<UploadDocVM>();

            foreach (var item in dirs)
            {
                string FolderName = System.IO.Path.GetFileName(item);
                UploadDocVM obj = new UploadDocVM()
                {
                    DirName = FolderName
                };
                lst.Add(obj);
            }

            ViewBag.LstCmp = new SelectList(_cmpservices.GetAll(), "id", "name");

            return View(lst);
        }

        [HttpPost]

        public IActionResult UploadDocuments(UploadDocVM obj)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            StringBuilder sb = new StringBuilder();
            try
            {


                if (obj.files != null && obj.files.Count > 0)
                {
                    sb.Append("<ul>");
                    foreach (var file in obj.files)
                    {
                        // Get the file name from the browser


                        var fname = Path.GetFileNameWithoutExtension(file.FileName);

                        var fulllname = fname.Split('_')[0];

                        var lastname = fulllname.Split(',')[0].TrimStart();

                        var firstname = "";



                        if (fulllname.Split(',').Length > 1)
                            firstname = fulllname.Split(',')[1].TrimStart();


                        var patientData = _patientservices.GetAll(" and fname='" + firstname + "' and lname='" + lastname + "' and cmp_id=" + cmpid);


                        if (patientData != null && patientData.Count == 1)
                        {
                            string PatientID = patientData[0].id.ToString();

                            var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", obj.DirName, PatientID);


                            bool folderExists = Directory.Exists(FolderPath);
                            if (!folderExists)
                                Directory.CreateDirectory(FolderPath);

                            var filePath = Path.Combine(FolderPath, file.FileName);

                            // Check If file with same name exists and delete it
                            if (System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Delete(filePath);
                            }

                            // Create a new local file and copy contents of uploaded file
                            using (var localFile = System.IO.File.OpenWrite(filePath))
                            using (var uploadedFile = file.OpenReadStream())
                            {
                                uploadedFile.CopyTo(localFile);
                            }
                            sb.AppendFormat("<li style='color:green'>{0} - Uploaded Successfully</li>", file.FileName);
                        }
                        else if (patientData != null && patientData.Count > 1)
                        {

                            //var match = Regex.Match(fname, @"_C(\d{1,2})_");

                            var match = Regex.Match(
    fname,
    @"(?:\(|_)\s*(?:Case|C)\s*(\d+)\s*\)?",
    RegexOptions.IgnoreCase
);

                            if (match.Success)
                            {
                                string accountNo = match.Groups[0].Value.TrimStart('_');
                                patientData = _patientservices.GetAll(" and fname='" + firstname + "' and lname='" + lastname + "' and account_no='" + accountNo + "' and cmp_id=" + cmpid);

                                if (patientData != null)
                                {
                                    string PatientID = patientData[0].id.ToString();

                                    var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", obj.DirName, PatientID);


                                    bool folderExists = Directory.Exists(FolderPath);
                                    if (!folderExists)
                                        Directory.CreateDirectory(FolderPath);

                                    var filePath = Path.Combine(FolderPath, file.FileName);

                                    // Check If file with same name exists and delete it
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        System.IO.File.Delete(filePath);
                                    }

                                    // Create a new local file and copy contents of uploaded file
                                    using (var localFile = System.IO.File.OpenWrite(filePath))
                                    using (var uploadedFile = file.OpenReadStream())
                                    {
                                        uploadedFile.CopyTo(localFile);
                                    }
                                    sb.AppendFormat("<li style='color:green'>{0} - Uploaded Successfully</li>", file.FileName);
                                }
                                else
                                {
                                    sb.AppendFormat("<li style='color:red'>{0} - No Patient found</li>", file.FileName);
                                }
                            }
                            else
                            {
                                patientData = _patientservices.GetAll(" and fname='" + firstname + "' and lname='" + lastname + "' and cmp_id=" + cmpid);

                                if (patientData != null)
                                {
                                    string PatientID = patientData[0].id.ToString();

                                    var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", obj.DirName, PatientID);


                                    bool folderExists = Directory.Exists(FolderPath);
                                    if (!folderExists)
                                        Directory.CreateDirectory(FolderPath);

                                    var filePath = Path.Combine(FolderPath, file.FileName);

                                    // Check If file with same name exists and delete it
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        System.IO.File.Delete(filePath);
                                    }

                                    // Create a new local file and copy contents of uploaded file
                                    using (var localFile = System.IO.File.OpenWrite(filePath))
                                    using (var uploadedFile = file.OpenReadStream())
                                    {
                                        uploadedFile.CopyTo(localFile);
                                    }
                                    sb.AppendFormat("<li style='color:green'>{0} - Uploaded Successfully</li>", file.FileName);
                                }
                                else
                                {
                                    sb.AppendFormat("<li style='color:red'>{0} - No Patient found</li>", file.FileName);
                                }
                            }
                        }
                        else
                        {
                            sb.AppendFormat("<li style='color:red'>{0} - No Patient found</li>", file.FileName);
                        }
                    }
                    sb.Append("</ul>");

                }
            }
            catch (Exception ex)
            {

                return Json(false);
            }

            return Json(sb.ToString());
        }

        public IActionResult UploadPatient()
        {
            return View(new UploadSignVM());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UploadPatient(IFormFile patient)
        {
            ViewBag.Message = "";

            try
            {
                int cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).Value;

                if (patient != null && patient.Length > 0)
                {
                    //DataTable dt = this.Read2007Xlsx(patient);
                    DataTable dt = new DataTable();
                    using (var stream = new MemoryStream())
                    {
                        patient.CopyToAsync(stream);
                        stream.Position = 0;

                        // Convert uploaded Excel to DataTable
                        dt = ReadExcelToDataTable(stream);
                    }

                    if (dt != null && dt.Rows.Count > 0)
                    {

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //tbl_ie_page1 obj = new tbl_ie_page1()
                            //{
                            //    cmp_id = cmpid,
                            //    allergies = dt.Rows[i]["Allergies"].ToString(),
                            //    cc = dt.Rows[i]["CC"].ToString(),
                            //    history = dt.Rows[i]["History"].ToString(),
                            //    ie_id = Convert.ToInt32(dt.Rows[i]["Patient_ie_id"].ToString()),
                            //    medication = dt.Rows[i]["Medications"].ToString(),
                            //    patient_id = Convert.ToInt32(dt.Rows[i]["Patient_id"].ToString()),
                            //    pe = dt.Rows[i]["Physical Exam"].ToString(),
                            //    pmh = dt.Rows[i]["Past Medical"].ToString(),
                            //    psh = dt.Rows[i]["Past Surgical"].ToString(),
                            //    social_history = dt.Rows[i]["Social History"].ToString(),
                            //    assessment = this.getAssement(dt.Rows[i]["Diagnoses"].ToString()),
                            //    //appt_reason = dt.Rows[i]["Reason"].ToString(),
                            //    //occupation = dt.Rows[i]["Occupation"].ToString(),
                            //    //impairment_rating = dt.Rows[i]["Tylenol"].ToString(),

                            //};
                            tbl_ie_page1 obj = new tbl_ie_page1()
                            {
                                cmp_id = cmpid,
                                allergies = dt.Columns.Contains("Allergies") ? dt.Rows[i]["Allergies"]?.ToString() ?? string.Empty : string.Empty,
                                cc = dt.Columns.Contains("CC") ? dt.Rows[i]["CC"]?.ToString() ?? string.Empty : string.Empty,
                                history = dt.Columns.Contains("History") ? dt.Rows[i]["History"]?.ToString() ?? string.Empty : string.Empty,
                                ie_id = dt.Columns.Contains("Patient_ie_id") ? Convert.ToInt32(dt.Rows[i]["Patient_ie_id"] ?? 0) : 0,
                                medication = dt.Columns.Contains("Medications") ? dt.Rows[i]["Medications"]?.ToString() ?? string.Empty : string.Empty,
                                patient_id = dt.Columns.Contains("Patient_id") ? Convert.ToInt32(dt.Rows[i]["Patient_id"] ?? 0) : 0,
                                pe = dt.Columns.Contains("Physical Exam") ? dt.Rows[i]["Physical Exam"]?.ToString() ?? string.Empty : string.Empty,
                                pmh = dt.Columns.Contains("Past Medical") ? dt.Rows[i]["Past Medical"]?.ToString() ?? string.Empty : string.Empty,
                                psh = dt.Columns.Contains("Past Surgical") ? dt.Rows[i]["Past Surgical"]?.ToString() ?? string.Empty : string.Empty,
                                social_history = dt.Columns.Contains("Social History") ? dt.Rows[i]["Social History"]?.ToString() ?? string.Empty : string.Empty,
                                assessment = dt.Columns.Contains("Diagnoses") ? this.getAssement(dt.Rows[i]["Diagnoses"]?.ToString() ?? string.Empty) : string.Empty,
                                appt_reason = dt.Columns.Contains("Reason") ? dt.Rows[i]["Reason"]?.ToString() ?? string.Empty : string.Empty,
                                occupation = dt.Columns.Contains("Occupation") ? dt.Rows[i]["Occupation"]?.ToString() ?? string.Empty : string.Empty,
                                impairment_rating = dt.Columns.Contains("Tylenol") ? dt.Rows[i]["Tylenol"]?.ToString() ?? string.Empty : string.Empty,
                            };
                            _patientimportservice.InsertPage1(obj);

                            tbl_ie_page2 obj2 = new tbl_ie_page2()
                            {
                                aod = dt.Columns.Contains("Activities") ? dt.Rows[i]["Activities"]?.ToString() ?? string.Empty : string.Empty,
                                ros = dt.Columns.Contains("ROS") ? dt.Rows[i]["ROS"]?.ToString() ?? string.Empty : string.Empty,
                                ie_id = Convert.ToInt32(dt.Rows[i]["Patient_ie_id"].ToString()),
                                cmp_id = cmpid,
                                patient_id = Convert.ToInt32(dt.Rows[i]["Patient_id"].ToString()),

                            };

                            _patientimportservice.InsertPage2(obj2);


                            var objNE = _patientimportservice.GetOneNE(Convert.ToInt32(dt.Rows[i]["Patient_ie_id"].ToString()));

                            if (objNE == null)
                            {
                                objNE = new tbl_ie_ne();

                                objNE.cmp_id = cmpid;
                                objNE.ie_id = dt.Columns.Contains("Patient_ie_id") ? Convert.ToInt32(dt.Rows[i]["Patient_ie_id"] ?? 0) : 0;
                                objNE.manual_muscle_strength_testing = dt.Columns.Contains("ManualMuscle") ? dt.Rows[i]["ManualMuscle"]?.ToString() ?? string.Empty : string.Empty;

                                if (string.IsNullOrEmpty(objNE.manual_muscle_strength_testing))
                                    objNE.manual_muscle_strength_testing = dt.Columns.Contains("Motor") ? dt.Rows[i]["Motor"]?.ToString() ?? string.Empty : string.Empty;

                                objNE.neurological_exam = dt.Columns.Contains("Neurological") ? dt.Rows[i]["Neurological"]?.ToString() ?? string.Empty : string.Empty;
                                objNE.sensory = dt.Columns.Contains("Sensory") ? dt.Rows[i]["Sensory"]?.ToString() ?? string.Empty : string.Empty;
                                objNE.patient_id = dt.Columns.Contains("Patient_id") ? Convert.ToInt32(dt.Rows[i]["Patient_id"] ?? 0) : 0;


                                objNE.other_content = dt.Columns.Contains("DeepTendon") ? dt.Rows[i]["DeepTendon"]?.ToString() ?? string.Empty : string.Empty;

                                if (string.IsNullOrEmpty(objNE.other_content))
                                {
                                    if (dt.Columns.Contains("Reflexes"))
                                        objNE.other_content = dt.Columns.Contains("Reflexes") ? dt.Rows[i]["Reflexes"]?.ToString() ?? string.Empty : string.Empty;
                                    else if (dt.Columns.Contains("Reflex"))
                                        objNE.other_content = dt.Columns.Contains("Reflex") ? dt.Rows[i]["Reflex"]?.ToString() ?? string.Empty : string.Empty;
                                }

                                _patientimportservice.InsertNE(objNE);
                            }
                            else
                            {
                                var _objNE = new tbl_ie_ne();


                                _objNE.manual_muscle_strength_testing = dt.Columns.Contains("ManualMuscle") ? dt.Rows[i]["ManualMuscle"]?.ToString() ?? string.Empty : string.Empty;
                                if (string.IsNullOrEmpty(_objNE.manual_muscle_strength_testing))
                                    _objNE.manual_muscle_strength_testing = dt.Columns.Contains("Motor") ? dt.Rows[i]["Motor"]?.ToString() ?? string.Empty : string.Empty;
                                _objNE.neurological_exam = dt.Columns.Contains("Neurological") ? dt.Rows[i]["Neurological"]?.ToString() ?? string.Empty : string.Empty;
                                _objNE.sensory = dt.Columns.Contains("Sensory") ? dt.Rows[i]["Sensory"]?.ToString() ?? string.Empty : string.Empty;
                                _objNE.id = objNE.id;
                                _objNE.other_content = dt.Columns.Contains("DeepTendon") ? dt.Rows[i]["DeepTendon"]?.ToString() ?? string.Empty : string.Empty;
                                if (string.IsNullOrEmpty(_objNE.other_content))
                                    _objNE.other_content = dt.Columns.Contains("Reflexes") ? dt.Rows[i]["Reflexes"]?.ToString() ?? string.Empty : string.Empty;



                                _patientimportservice.UpdatePageNE(_objNE);
                            }

                        }

                        ViewBag.Message = "Patient Data uploaded successfully.";
                    }

                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = "";
                SaveLog(ex, "ImportData");
            }
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UploadPatientFU(IFormFile patientFU)
        {
            ViewBag.Message = "";
            try
            {
                int cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).Value;

                if (patientFU != null && patientFU.Length > 0)
                {
                    //DataTable dt = this.Read2007Xlsx(patient);
                    DataTable dt = new DataTable();
                    using (var stream = new MemoryStream())
                    {
                        patientFU.CopyToAsync(stream);
                        stream.Position = 0;

                        // Convert uploaded Excel to DataTable
                        dt = ReadExcelToDataTable(stream);
                    }

                    if (dt != null && dt.Rows.Count > 0)
                    {

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //tbl_fu_page1 obj = new tbl_fu_page1()
                            //{
                            //    cmp_id = cmpid,
                            //    allergies = dt.Rows[i]["Allergies"].ToString(),
                            //    cc = dt.Rows[i]["CC"].ToString(),
                            //    //history = dt.Rows[i]["History"].ToString(),
                            //    ie_id = Convert.ToInt32(dt.Rows[i]["Patient_ie_id"].ToString()),
                            //    medication = dt.Rows[i]["Medications"].ToString(),
                            //    patient_id = Convert.ToInt32(dt.Rows[i]["Patient_id"].ToString()),
                            //    pe = dt.Rows[i]["Physical Exam"].ToString(),
                            //    pmh = dt.Rows[i]["Past Medical"].ToString(),
                            //    psh = dt.Rows[i]["Past Surgical"].ToString(),
                            //    social_history = dt.Rows[i]["Social History"].ToString(),
                            //    assessment = this.getAssement(dt.Rows[i]["Diagnoses"].ToString()),
                            //    appt_reason = dt.Rows[i]["Reason"].ToString(),
                            //    //occupation = dt.Rows[i]["Occupation"].ToString(),
                            //    //impairment_rating = dt.Rows[i]["Tylenol"].ToString(),
                            //    fu_id = Convert.ToInt32(dt.Rows[i]["Patient_fu_id"].ToString()),

                            //};

                            tbl_fu_page1 obj = new tbl_fu_page1()
                            {
                                cmp_id = cmpid,
                                allergies = dt.Columns.Contains("Allergies") ? dt.Rows[i]["Allergies"]?.ToString() ?? string.Empty : string.Empty,
                                cc = dt.Columns.Contains("CC") ? dt.Rows[i]["CC"]?.ToString() ?? string.Empty : string.Empty,
                                history = dt.Columns.Contains("History") ? dt.Rows[i]["History"]?.ToString() ?? string.Empty : string.Empty,
                                ie_id = dt.Columns.Contains("Patient_ie_id") ? Convert.ToInt32(dt.Rows[i]["Patient_ie_id"] ?? 0) : 0,
                                medication = dt.Columns.Contains("Medications") ? dt.Rows[i]["Medications"]?.ToString() ?? string.Empty : string.Empty,
                                patient_id = dt.Columns.Contains("Patient_id") ? Convert.ToInt32(dt.Rows[i]["Patient_id"] ?? 0) : 0,
                                pe = dt.Columns.Contains("Physical Exam") ? dt.Rows[i]["Physical Exam"]?.ToString() ?? string.Empty : string.Empty,
                                pmh = dt.Columns.Contains("Past Medical") ? dt.Rows[i]["Past Medical"]?.ToString() ?? string.Empty : string.Empty,
                                psh = dt.Columns.Contains("Past Surgical") ? dt.Rows[i]["Past Surgical"]?.ToString() ?? string.Empty : string.Empty,
                                social_history = dt.Columns.Contains("Social History") ? dt.Rows[i]["Social History"]?.ToString() ?? string.Empty : string.Empty,
                                assessment = dt.Columns.Contains("Diagnoses") ? this.getAssement(dt.Rows[i]["Diagnoses"]?.ToString() ?? string.Empty) : string.Empty,
                                appt_reason = dt.Columns.Contains("Reason") ? dt.Rows[i]["Reason"]?.ToString() ?? string.Empty : string.Empty,
                                occupation = dt.Columns.Contains("Occupation") ? dt.Rows[i]["Occupation"]?.ToString() ?? string.Empty : string.Empty,
                                impairment_rating = dt.Columns.Contains("Tylenol") ? dt.Rows[i]["Tylenol"]?.ToString() ?? string.Empty : string.Empty,
                                fu_id = dt.Columns.Contains("Patient_fu_id") ? Convert.ToInt32(dt.Rows[i]["Patient_fu_id"] ?? 0) : 0
                            };
                            _patientimportservice.InsertPage1FU(obj);

                            tbl_fu_page2 obj2 = new tbl_fu_page2()
                            {
                                aod = dt.Columns.Contains("Activities") ? dt.Rows[i]["Activities"]?.ToString() ?? string.Empty : string.Empty,
                                ros = dt.Columns.Contains("ROS") ? dt.Rows[i]["ROS"]?.ToString() ?? string.Empty : string.Empty,
                                ie_id = Convert.ToInt32(dt.Rows[i]["Patient_ie_id"].ToString()),
                                cmp_id = cmpid,
                                patient_id = Convert.ToInt32(dt.Rows[i]["Patient_id"].ToString()),
                                fu_id = Convert.ToInt32(dt.Rows[i]["Patient_fu_id"].ToString()),
                            };

                            _patientimportservice.InsertPage2FU(obj2);

                            //if (dt.Columns.Contains("Neurological"))
                            //{
                            var objNE = _patientimportservice.GetOneNEFU(Convert.ToInt32(dt.Rows[i]["Patient_fu_id"].ToString()));

                            if (objNE == null)
                            {
                                objNE = new tbl_fu_ne();

                                objNE.cmp_id = cmpid;
                                objNE.ie_id = dt.Columns.Contains("Patient_ie_id") ? Convert.ToInt32(dt.Rows[i]["Patient_ie_id"] ?? 0) : 0;
                                objNE.manual_muscle_strength_testing = dt.Columns.Contains("ManualMuscle") ? dt.Rows[i]["ManualMuscle"]?.ToString() ?? string.Empty : string.Empty;

                                if (string.IsNullOrEmpty(objNE.manual_muscle_strength_testing))
                                    objNE.manual_muscle_strength_testing = dt.Columns.Contains("Motor") ? dt.Rows[i]["Motor"]?.ToString() ?? string.Empty : string.Empty;



                                objNE.neurological_exam = dt.Columns.Contains("Neurological") ? dt.Rows[i]["Neurological"]?.ToString() ?? string.Empty : string.Empty;
                                objNE.sensory = dt.Columns.Contains("Sensory") ? dt.Rows[i]["Sensory"]?.ToString() ?? string.Empty : string.Empty;
                                objNE.patient_id = dt.Columns.Contains("Patient_id") ? Convert.ToInt32(dt.Rows[i]["Patient_id"] ?? 0) : 0;
                                objNE.other_content = dt.Columns.Contains("DeepTendon") ? dt.Rows[i]["DeepTendon"]?.ToString() ?? string.Empty : string.Empty;

                                if (string.IsNullOrEmpty(objNE.other_content))
                                {
                                    if (dt.Columns.Contains("Reflexes"))
                                        objNE.other_content = dt.Columns.Contains("Reflexes") ? dt.Rows[i]["Reflexes"]?.ToString() ?? string.Empty : string.Empty;
                                    else if (dt.Columns.Contains("Reflex"))
                                        objNE.other_content = dt.Columns.Contains("Reflex") ? dt.Rows[i]["Reflex"]?.ToString() ?? string.Empty : string.Empty;
                                }

                                objNE.fu_id = dt.Columns.Contains("Patient_fu_id") ? Convert.ToInt32(dt.Rows[i]["Patient_fu_id"] ?? 0) : 0;


                                _patientimportservice.InsertNEFU(objNE);
                            }
                            else
                            {
                                var _objNE = new tbl_fu_ne();


                                _objNE.manual_muscle_strength_testing = dt.Columns.Contains("ManualMuscle") ? dt.Rows[i]["ManualMuscle"]?.ToString() ?? string.Empty : string.Empty;

                                if (string.IsNullOrEmpty(objNE.manual_muscle_strength_testing))
                                    _objNE.manual_muscle_strength_testing = dt.Columns.Contains("Motor") ? dt.Rows[i]["Motor"]?.ToString() ?? string.Empty : string.Empty;



                                _objNE.neurological_exam = dt.Columns.Contains("Neurological") ? dt.Rows[i]["Neurological"]?.ToString() ?? string.Empty : string.Empty;
                                _objNE.sensory = dt.Columns.Contains("Sensory") ? dt.Rows[i]["Sensory"]?.ToString() ?? string.Empty : string.Empty;
                                _objNE.other_content = dt.Columns.Contains("DeepTendon") ? dt.Rows[i]["DeepTendon"]?.ToString() ?? string.Empty : string.Empty;

                                if (string.IsNullOrEmpty(_objNE.other_content))
                                    _objNE.other_content = dt.Columns.Contains("Reflexes") ? dt.Rows[i]["Reflexes"]?.ToString() ?? string.Empty : string.Empty;


                                _objNE.id = objNE.id;


                                _patientimportservice.UpdateNEFU(_objNE);
                            }
                            //}
                        }
                        ViewBag.Message = "Patient Data uploaded successfully.";
                    }

                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "ImportData");
                ViewBag.Message = "";
            }
            return View("UploadPatient");
        }

        private string getAssement(string val)
        {
            if (string.IsNullOrEmpty(val))
                return "";
            else
            {
                val = val.Replace("<p>", "<li>");
                val = val.Replace("</p>", "</li>");
                val = "<ul>" + val + "</ul>";
                return val;
            }
        }

        public DataTable Read2007Xlsx(IFormFile postedFile)
        {
            DataTable dt = new DataTable();
            try
            {
                if (postedFile != null && postedFile.Length > 0)
                {
                    // Read the uploaded Excel file using Open XML
                    using (Stream stream = postedFile.OpenReadStream())
                    {

                        using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(stream, false))
                        {
                            WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                            IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                            }
                            foreach (Row row in rows) //this will also include your header row...
                            {
                                DataRow tempRow = dt.NewRow();
                                int columnIndex = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    // Gets the column index of the cell with data
                                    int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                                    cellColumnIndex--; //zero based index
                                    if (columnIndex < cellColumnIndex)
                                    {
                                        do
                                        {
                                            tempRow[columnIndex] = ""; //Insert blank data here;
                                            columnIndex++;
                                        }
                                        while (columnIndex < cellColumnIndex);
                                    }//end if block
                                    tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);
                                    columnIndex++;
                                }//end inner foreach loop
                                dt.Rows.Add(tempRow);
                            }//end outer foreach loop
                        }//end using block
                        dt.Rows.RemoveAt(0); //...so i'm taking it out here.
                    }
                }
            }//end try
            catch (Exception ex)
            {
                SaveLog(ex, "Read2007Xlsx");
            }

            return dt;
        }//end Read2007Xlsx method        

        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        } //end GetColumnName method
        public static int? GetColumnIndexFromName(string columnName)
        {
            //return columnIndex;
            string name = columnName;
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        } //end GetColumnIndexFromName method

        //public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        //{
        //    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
        //    if (cell.CellValue == null)
        //    {
        //        return "";
        //    }
        //    string value = cell.CellValue.InnerXml;
        //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //    {
        //        return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
        //    }
        //    else
        //    {
        //        return value;
        //    }
        //}//end GetCellValue method


        private DataTable ReadExcelToDataTable(Stream stream)
        {
            DataTable dataTable = new DataTable();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheets sheets = workbookPart.Workbook.Sheets;

                // Get the first sheet
                Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault();

                if (sheet == null)
                {
                    throw new Exception("No sheet found in the Excel file.");
                }

                // Get the WorksheetPart based on the sheet's relationship ID
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                bool isFirstRow = true;
                foreach (Row row in sheetData.Elements<Row>())
                {
                    DataRow dataRow = dataTable.NewRow();

                    int columnIndex = 0;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(spreadsheetDocument, cell);

                        if (isFirstRow)
                        {
                            // Use the first row to add columns to the DataTable
                            dataTable.Columns.Add(cellValue);
                        }
                        else
                        {
                            // Clean HTML tags and add data to the DataTable
                            //string cleanCellValue = System.Text.RegularExpressions.Regex.Replace(cellValue, "<.*?>", string.Empty);
                            string cleanCellValue = cellValue;
                            dataRow[columnIndex] = cleanCellValue;
                        }
                        columnIndex++;
                    }

                    if (!isFirstRow)
                    {
                        dataTable.Rows.Add(dataRow);
                    }
                    isFirstRow = false;
                }
            }

            return dataTable;
        }

        private string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell == null)
                return null;

            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                value = stringTable.ChildElements[int.Parse(value)].InnerText;
            }

            return value;
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
