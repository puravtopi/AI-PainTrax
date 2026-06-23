using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Models;
using MS.Services;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;

namespace PainTrax.Web.Controllers
{
    public class ImportCSVController : Controller
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

        public ImportCSVController(ILogger<FormsController> logger, IWebHostEnvironment environment)
        {
            _environment = environment;
            _logger = logger;
        }

        private int calculateAge(DateTime bday)
        {
            DateTime today = DateTime.Today;

            int age = today.Year - bday.Year;

            if (today.Month < bday.Month || ((today.Month == bday.Month) && (today.Day < bday.Day)))
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


     

        public int GetIns(Dictionary<string, string> row, int? cmpid, int? userid)
        {
            int priminsId = 0;
            string query = "";

            List<tbl_inscos> insdata = new List<tbl_inscos>();

            tbl_inscos objInscos = new tbl_inscos();

            string insurance = row.ContainsKey("Insurance Company") ? row["Insurance Company"] : "";

            if (!string.IsNullOrEmpty(insurance))
            {
                query = " and cmpname='" + insurance +
                        "' and cmp_id=" + cmpid;

                insdata = _inscosservices.GetAll(query);

                objInscos = new tbl_inscos()
                {
                    cmpname = insurance,
                    cmp_id = cmpid,
                    createdby = userid,
                    createddate = DateTime.Now
                };

                if (insdata.Count > 0)
                {
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
            int priminsId = 0;
            string query = "";

            List<tbl_inscos> insdata = new List<tbl_inscos>();

            tbl_inscos objInscos = new tbl_inscos();

            string insurance = row.ContainsKey("Secondary Insurance") ? row["Secondary Insurance"] : "";

            if (!string.IsNullOrEmpty(insurance))
            {
                query = " and cmpname='" + insurance +
                        "' and cmp_id=" + cmpid;

                insdata = _inscosservices.GetAll(query);

                objInscos = new tbl_inscos()
                {
                    cmpname = insurance,
                    cmp_id = cmpid,
                    createdby = userid,
                    createddate = DateTime.Now
                };

                if (insdata.Count > 0)
                {
                    priminsId = insdata[0].id.Value;
                }
                else
                {
                    priminsId = _inscosservices.Insert(objInscos);
                }
            }


            return priminsId;
        }



        private List<Dictionary<string, string>> ExtractKeyValuePairs(IFormFile file)
        {
            var records = new List<Dictionary<string, string>>();

            if (file == null || file.Length == 0)
                return records;

            using (var stream = file.OpenReadStream())
            using (var reader = new StreamReader(stream))
            {
                // Read Header
                var headerLine = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(headerLine))
                    return records;

                var headers = headerLine.Split(',')
                                        .Select(h => h.Trim())
                                        .ToList();

                // Read Data
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();

                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var values = line.Split(',');

                    var rowData = new Dictionary<string, string>();

                    for (int i = 0; i < headers.Count; i++)
                    {
                        rowData[headers[i]] = i < values.Length
                            ? values[i].Trim()
                            : "";
                    }

                    records.Add(rowData);
                }
            }

            return records;
        }
        public List<(string PatientName, string Message, bool IsInserted)> SaveDetails(List<Dictionary<string, string>> data, string locationid)
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
                            DateTime.ParseExact(row["Date of Birth"], "M/d/yyyy", null)
                        );

                        // var nameValue = row["Name"];

                        tbl_patient objPatient = new tbl_patient()
                        {

                            fname = row.ContainsKey("First Name") ? row["First Name"] : "",
                            lname = row.ContainsKey("Last Name") ? row["Last Name"] : "",
                            gender = row["Gender"] == "Male" ? "1" : row["Gender"] == "Female" ? "2" : "3",
                            address = (row.ContainsKey("Street Address")? row["Street Address"]: ""),
                            dob = DateTime.ParseExact(row["Date of Birth"], "M/d/yyyy", null),
                            email = row.ContainsKey("Email") ? row["Email"] : "",
                            home_ph = row.ContainsKey("Home Phone") ? row["Home Phone"] : "",
                            
                            mobile = row.ContainsKey("Cell Phone") ? row["Cell Phone"] : "",
                            ssn = row.ContainsKey("Social Security Number") ? row["Social Security Number"] : "",
                            city = row.ContainsKey("City") ? row["City"] : "",
                            state = row.ContainsKey("State") ? row["State"] : "",
                            zip = row.ContainsKey("Zip") ? row["Zip"] : "",
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
                        //empId = GetEmployer(row, patientId);
                        //attornyId = GetAttorney(row, cmpid, userid);
                        int? providerid = null;
                        DataTable prov_data = null;
                        if (row.ContainsKey("Doctor") && !string.IsNullOrEmpty( row["Doctor"]))
                        {
                            var prov_query = $"SELECT id FROM tbl_users WHERE desigid in (SELECT id FROM tbl_designation WHERE title='Provider' AND cmp_id ={cmpid} ) AND fullname='{row["Doctor"]}' ";

                            prov_data = _pareentservices.GetData(prov_query);
                            if (prov_data != null && prov_data.Rows.Count > 0)
                                providerid = Convert.ToInt32( prov_data.Rows[0]["id"].ToString());
                        }



                        // Save IE Details
                        tbl_patient_ie objIE = new tbl_patient_ie()
                        {
                            created_by = userid,
                            created_date = DateTime.Now,
                            doe = DateTime.ParseExact(row["Last Visit Date"], "M/d/yyyy", null),
                            //doa = DateTime.ParseExact(row["Accident Date"], "MM/dd/yyyy", null),
                            emp_id = empId,
                            is_active = true,
                            location_id = Convert.ToInt32(locationid),
                            patient_id = patientId,
                            provider_id = providerid
                            // primary_ins_cmp_id = priminsId,
                            //secondary_ins_cmp_id = secinsId,
                            //primary_policy_no = row.ContainsKey("Policy#") ? row["Policy#"] : "",
                            //secondary_policy_no = row.ContainsKey("Sec Policy#") ? row["Sec Policy#"] : "",
                            //primary_claim_no = row.ContainsKey("Claim#") ? row["Claim#"] : "",
                            //secondary_claim_no = row.ContainsKey("Sec Claim#") ? row["Sec Claim#"] : "",
                            // primary_wcb_group = row.ContainsKey("Wcb#") ? row["Wcb#"] : "",
                            /* compensation = row.ContainsKey("Case Type")
                                 ? (
                                     row["Case Type"] == "PVT" ? "PI" :
                                     row["Case Type"] == "NF" ? "NF" :
                                     row["Case Type"] == "WC" ? "WC" :
                                     row["Case Type"] == "Lien" ? "Lien" :
                                     row["Case Type"] == "PI" ? "PI" :
                                     ""
                                   )
                                 : "",*/

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
            if (locationid == "0")
            {
                TempData["Message"] = "Please Select Location";
            }
            else
            {
                if (files != null)
                {
                    string message = "";
                    foreach (var file in files)
                    {
                        var csvdata = ExtractKeyValuePairs(file);
                        var result = SaveDetails(csvdata, locationid);
                        //foreach (var row in csvdata)
                        //{
                        //    message += "\n Record:";

                        //    foreach (var item in row)
                        //    {
                        //        message += $"{item.Key} : {item.Value}";
                        //    }


                        //}
                        foreach (var item in result)
                        {
                            message += $"{item.PatientName} | {item.Message}  <br>";
                        }

                    }
                    TempData["Message"] = message;
                }
                else
                {
                    TempData["Message"] = "File Not Uploaded.";
                }
            }
            return RedirectToAction("Index");
        }
    }
}
