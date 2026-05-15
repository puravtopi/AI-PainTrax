using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Models;
using MS.Services;
using OfficeOpenXml;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Collections.Generic;

namespace PainTrax.Web.Controllers
{
    public class ImportExcelController : Controller
    {
        private readonly ILogger<FormsController> _logger;

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

        public ImportExcelController(ILogger<FormsController> logger, IWebHostEnvironment environment)
        {
            _environment = environment;
            _logger = logger;
        }

        private int calculateAge(DateTime bday)
        {
            DateTime today = DateTime.Today;

            int age = today.Year - bday.Year;

            if (today.Month < bday.Month ||   ((today.Month == bday.Month) && (today.Day < bday.Day)))
            {
                age--;  //birthday in current year not yet reached, we are 1 year younger ;)
                        //+ no birthday for 29.2. guys ... sorry, just wrong date for birth
            }

            return age;
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
            return View();
        }

        private List<Dictionary<string, string>> ExtractKeyValuePairs(IFormFile file)
        {
            var records = new List<Dictionary<string, string>>();

            if (file == null || file.Length == 0)
                return records;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);

                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    if (worksheet == null)
                        return records;

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // First row = headers
                    List<string> headers = new List<string>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text.Trim());
                    }

                    // Data rows
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var rowData = new Dictionary<string, string>();

                        for (int col = 1; col <= colCount; col++)
                        {
                            string key = headers[col - 1];
                            string value = worksheet.Cells[row, col].Text.Trim();

                            rowData[key] = value;
                        }

                        records.Add(rowData);
                    }
                }
            }

            return records;
        }
        public int GetIns(Dictionary<string, string> row, int? cmpid,int? userid)
        {
            int priminsId = 0;
            string query = "";

            List<tbl_inscos> insdata = new List<tbl_inscos>();

            tbl_inscos objInscos = new tbl_inscos();

            string insurance = row.ContainsKey("Insurance") ? row["Insurance"] : "";

            if (!string.IsNullOrEmpty(insurance))
            {
                query = " and cmpname='" + insurance +
                        "' and cmp_id=" + cmpid;

                insdata = _inscosservices.GetAll(query);

                objInscos = new tbl_inscos()
                {
                    cmpname = insurance,
                    address1 = row.ContainsKey("Insurance Address 1") ? row["Insurance Address 1"] : "",
                    address2 = row.ContainsKey("Insurance Address 2") ? row["Insurance Address 2"] : "",
                    city = row.ContainsKey("Insurance City") ? row["Insurance City"] : "",       
                    state = row.ContainsKey("Insurance City") ? row["Insurance State"] : "",
                    zipcode  = row.ContainsKey("Insurance City") ? row["Insurance Zip"] : "",
                    telephone = row.ContainsKey("Insurance Phone") ? row["Insurance Phone"] : "",
                    faxno = row.ContainsKey("Insurance Phone") ? row["Insurance Fax"] : "",
                    cmp_id = cmpid,
                    createdby = userid,
                    createddate = DateTime.Now
                };

                if (insdata.Count > 0)
                {
                   // objInscos.id = insdata[0].id.Value;

                   // _inscosservices.Update(objInscos);

                    priminsId = insdata[0].id.Value;
                }
                else
                {
                    priminsId = _inscosservices.Insert(objInscos);
                }
            }


            return priminsId;
        }
        public int GetSecIns(Dictionary<string, string> row, int? cmpid, int? userid)
        {
            int secinsId = 0;
            string query = "";

            List<tbl_inscos> insdata = new List<tbl_inscos>();

            tbl_inscos objInscos = new tbl_inscos();

            string insurance = row.ContainsKey("Sec Insurance") ? row["Sec Insurance"] : "";

            if (!string.IsNullOrEmpty(insurance))
            {
                query = " and cmpname='" + insurance +
                        "' and cmp_id=" + cmpid;

                insdata = _inscosservices.GetAll(query);

                objInscos = new tbl_inscos()
                {
                    cmpname = insurance,
                   
                    address1 = row.ContainsKey("Sec Insurance Address 1") ? row["Sec Insurance Address 1"] : "",
                    address2 = row.ContainsKey("Sec Insurance Address 2") ? row["Sec Insurance Address 2"] : "",
                    city = row.ContainsKey("Sec Insurance City") ? row["Sec Insurance City"] : "",
                    state = row.ContainsKey("Sec Insurance City") ? row["Sec Insurance State"] : "",
                    zipcode = row.ContainsKey("Sec Insurance City") ? row["Sec Insurance Zip"] : "",
                    telephone = row.ContainsKey("Sec Insurance Phone") ? row["Sec Insurance Phone"] : "",
                    faxno = row.ContainsKey("Sec Insurance Phone") ? row["Sec Insurance Fax"] : "",
                    cmp_id = cmpid,
                    createdby = userid,
                    createddate = DateTime.Now
                };

                if (insdata.Count > 0)
                {
                   // objInscos.id = insdata[0].id.Value;

                    //_inscosservices.Update(objInscos);

                    secinsId = insdata[0].id.Value;
                }
                else
                {
                    secinsId = _inscosservices.Insert(objInscos);
                }
            }
            return secinsId;
        }

        public int GetEmployer(Dictionary<string, string> row, int patientId)
        {
            int empId = 0;

            string employerName = row.ContainsKey("Employer Name")
                ? row["Employer Name"]
                : "";

            if (!string.IsNullOrEmpty(employerName))
            {
                string address =
                    (row.ContainsKey("Empoyer Address")
                        ? row["Empoyer Address"]
                        : "")
                    + " " +
                    (row.ContainsKey("Employer City")
                        ? row["Employer City"]
                        : "")
                    + " " +
                    (row.ContainsKey("Employer state")
                        ? row["Employer state"]
                        : "")
                    + " " +
                    (row.ContainsKey("Employer Zip")
                        ? row["Employer Zip"]
                        : "");

                tbl_emp objEmp = new tbl_emp()
                {
                    name = employerName,
                    address = address.Trim(),
                    phone = row.ContainsKey("Empoyer Phone")
                        ? row["Empoyer Phone"]
                        : "",
                    fax = "",
                    patient_id = patientId
                };

                // Optional duplicate check
                string query = " and name='" + employerName +
                               "' and patient_id=" + patientId;

                var empdata = _empService.GetAll(query);

                if (empdata.Count > 0)
                {
                    //objEmp.id = empdata[0].id;

                    //_empService.Update(objEmp);

                    empId = empdata[0].id ?? 0;
                }
                else
                {
                    empId = _empService.Insert(objEmp);
                }
            }

            return empId;
        }

        public int GetAttorney(Dictionary<string, string> row, int? cmpid, int? userid)
        {
            int attorneyId = 0;

            string attorneyName = row.ContainsKey("Attorney Name")
                ? row["Attorney Name"]
                : "";

            if (!string.IsNullOrWhiteSpace(attorneyName))
            {
                string query = " and Attorney='" + attorneyName +
                               "' and cmp_id=" + cmpid;

                List<tbl_attorneys> attorneyData = _attorneyservices.GetAll(query);

                if (attorneyData.Count > 0)
                {
                    attorneyId = attorneyData[0].Id ?? 0;
                }
                else
                {
                    tbl_attorneys objAttorney = new tbl_attorneys()
                    {
                        Attorney = attorneyName,
                        CreatedDate = DateTime.Now,
                        CreatedBy = userid,
                        cmp_id = cmpid
                    };

                    attorneyId = _attorneyservices.Insert(objAttorney);
                }
            }

            return attorneyId;
        }

        public List<(string PatientName, string Message, bool IsInserted)> SaveDetails(List<Dictionary<string, string>> data,  string locationid)
        {
            var results = new List<(string PatientName, string Message, bool IsInserted)>();

            try
            {
                int patientId = 0, priminsId = 0, secinsId = 0, attornyId = 0, adjusterId = 0, empId = 0;

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

                foreach (var row in data)
                {
                    string firstName = row.ContainsKey("First Name")
                        ? row["First Name"]
                        : "";

                    string lastName = row.ContainsKey("Last Name")
                        ? row["Last Name"]
                        : "";

                    string patientName = lastName + " " + firstName;

                    try
                    {
                        int age = calculateAge(
                            DateTime.ParseExact(row["DateOfBirth"], "MM/dd/yyyy", null)
                        );

                        // var nameValue = row["Name"];

                        tbl_patient objPatient = new tbl_patient()
                        {

                            fname = row.ContainsKey("First Name") ? row["First Name"] : "",
                            mname = row.ContainsKey("Last Name") ? row["Middle Name"] : "",
                            lname = row.ContainsKey("Last Name") ? row["Last Name"] : "",
                            gender = row["Gender"] == "Male" ? "1" : row["Gender"] == "Female" ? "2" : "3",
                            address = (row.ContainsKey("Patient Address Line 1")
                                        ? row["Patient Address Line 1"]
                                        : "")
                                    + " " +
                                    (row.ContainsKey("Patient Address Line 2")
                                        ? row["Patient Address Line 2"]
                                        : ""),
                            dob = DateTime.ParseExact(row["DateOfBirth"], "MM/dd/yyyy", null),
                            email = row.ContainsKey("Patient Email") ? row["Patient Email"] : "",
                            home_ph = row.ContainsKey("Home Phone") ? row["Home Phone"] : "",
                            mobile = row.ContainsKey("Cell Phone") ? row["Cell Phone"] : "",
                            ssn = row.ContainsKey("SSN") ? row["SSN"] : "",
                            state = row.ContainsKey("Patient State") ? row["Patient State"] : "",
                            zip = row.ContainsKey("Patient Zip") ? row["Patient Zip"] : "",
                            cmp_id = cmpid,
                            createdby = userid,
                            createddate = DateTime.Now,
                            age = age

                        };

                        List<tbl_patient> ptdata = new List<tbl_patient>();

                        var dobStr = objPatient.dob.HasValue
                            ? objPatient.dob.Value.ToString("yyyy-MM-dd")
                            : "";

                        var query = " and fname='" + objPatient.fname +
                                    "' and lname='" + objPatient.lname +
                                    "' and dob='" + dobStr +
                                    "' and cmp_id=" + cmpid;

                        ptdata = _patientservices.GetAll(query);

                        if (ptdata.Count > 0)
                        {
                            results.Add((
                                   patientName,
                                   "Already Exist",
                                   false
                               ));
                            continue;
                        }
                        else
                        {
                            patientId = _patientservices.Insert(objPatient);
                        }

                        HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, patientId);

                        ViewBag.patientId = patientId;

                        // Insurance
                        priminsId = GetIns(row, cmpid, userid);
                        secinsId = GetSecIns(row, cmpid, userid);
                        empId = GetEmployer(row, patientId);
                        attornyId = GetAttorney(row, cmpid, userid);
                        // Save IE Details
                        tbl_patient_ie objIE = new tbl_patient_ie()
                        {
                            adjuster_id = adjusterId,
                            attorney_id = attornyId,
                            created_by = userid,
                            created_date = DateTime.Now,
                            doe = DateTime.ParseExact(row["Date Of First Treatement"], "MM/dd/yyyy", null),
                            doa = DateTime.ParseExact(row["Accident Date"], "MM/dd/yyyy", null),
                            emp_id = empId,
                            is_active = true,
                            location_id = Convert.ToInt32(locationid),
                            patient_id = patientId,
                            primary_ins_cmp_id = priminsId,
                            secondary_ins_cmp_id = secinsId,
                            primary_policy_no = row.ContainsKey("Policy#") ? row["Policy#"] : "",
                            secondary_policy_no = row.ContainsKey("Sec Policy#") ? row["Sec Policy#"] : "",
                            primary_claim_no = row.ContainsKey("Claim#") ? row["Claim#"] : "",
                            secondary_claim_no = row.ContainsKey("Sec Claim#") ? row["Sec Claim#"] : "",
                            primary_wcb_group = row.ContainsKey("Wcb#") ? row["Wcb#"] : "",
                            compensation = row.ContainsKey("Case Type")
                                ? (
                                    row["Case Type"] == "PVT" ? "PI" :
                                    row["Case Type"] == "NF" ? "NF" :
                                    row["Case Type"] == "WC" ? "WC" :
                                    row["Case Type"] == "Lien" ? "Lien" :
                                    row["Case Type"] == "PI" ? "PI" :
                                    ""
                                  )
                                : "",

                        };

                        int ie = _ieService.Insert(objIE);
                        results.Add((
                              patientName,
                              "Imported",
                              true
                          ));
                    }
                    catch (Exception ex)
                    {
                        results.Add((
                            patientName,
                            ex.Message,
                            false
                        ));
                    }
                }
               
            }
            catch (Exception ex)
            {
                results.Add((
                    "",
                    ex.Message,
                    false
                ));

            }
            return results;
        }

        [HttpPost]
        public ActionResult UploadFile(List<IFormFile> files, string locationid)
        {
            if (files != null)
            {
                string message = "";
                foreach (var file in files)
                {
                    var exceldata = ExtractKeyValuePairs(file);
                    var  result = SaveDetails(exceldata, locationid);
                    //foreach (var row in exceldata)
                    //{
                    //    message += "\n Record:";

                    //    foreach (var item in row)
                    //    {
                    //        message += $"{item.Key} : {item.Value}";
                    //    }

                        
                    //}
                    foreach (var item in result)
                    {
                        message += $"{item.PatientName} | {item.Message} | {item.IsInserted}\n";
                    }

                }
                TempData["Message"] = message;
            }
            else
            {
                TempData["Message"] = "File Not Uploaded.";
            }
            return RedirectToAction("Index");
        }
    }
}
