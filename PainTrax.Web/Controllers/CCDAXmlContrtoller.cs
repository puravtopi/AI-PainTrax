using Microsoft.AspNetCore.Mvc;
using MS.Services;
using Optivem.Framework.Core.Domain;
using PainTrax.Services;
using PainTrax.Web.Helper;

using System.Data;
using System.IO.Compression;
using System.Text;

namespace PainTrax.Web.Controllers
{
    public class CCDAXmlController : Controller
    {
        private readonly IWebHostEnvironment _environment;
        private readonly PatientService _patientservices = new PatientService();
        private readonly PatientIEService _ieService = new PatientIEService();


        public CCDAXmlController(IWebHostEnvironment environment) 
        { 
            _environment = environment;
        }
        public ActionResult GenerateXml(string searchtxt = "")
        {
           
            return View();


        }
        public IActionResult List()
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
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][name]"].FirstOrDefault();
                // Sort Column Direction ( asc ,desc)
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                // Search Value from (Search box)
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                //Paging Size (10,20,50,100)
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                string cnd = " and patient_id in (select id from tbl_patient where cmp_id=" + cmpid + ")";

                if (!string.IsNullOrEmpty(searchValue))
                    cnd = " and (fname like '%" + searchValue + "%' or lname  like '%" + searchValue + "%' or location  like '%" + searchValue + "%' or DATE_FORMAT(dob,\"%m/%d/%Y\") = '" + searchValue + "' or DATE_FORMAT(doe,\"%m/%d/%Y\") = '" + searchValue + "') ";

                var Data = _ieService.GetAll(cnd);

                //Sorting

                //Search


                //total number of rows count 
                recordsTotal = Data.Count();
                //Paging 
                var data = Data.Skip(skip).Take(pageSize).ToList();
                //Returning Json Data
                return Json(new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data });

            }
            catch (Exception)
            {
                throw;
            }

        }

        public ActionResult DownloadZip(DateTime fromDate, DateTime toDate)
        {
            byte[] zipbytes = null;
            ParentService service = new ParentService();
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string sql = $@"
                            SELECT
                                tbl_patient.*,
                                tbl_patient_ie.*,
                                tbl_ie_page1.*,
                                DATE_FORMAT(tbl_patient.dob, '%Y%m%d') AS birthtime,
                                CASE
                                    WHEN tbl_patient.gender = 1 THEN 'M'
                                    WHEN tbl_patient.gender = 2 THEN 'F'
                                    ELSE 'Other'
                                    END AS gendercode
                            FROM tbl_patient
                            INNER JOIN tbl_patient_ie
                                ON tbl_patient.id = tbl_patient_ie.patient_id
                            LEFT JOIN tbl_ie_page1
                                ON tbl_patient.id = tbl_ie_page1.patient_id
                            WHERE tbl_patient_ie.created_date
                            BETWEEN '{fromDate:yyyy-MM-dd}'
                            AND '{toDate:yyyy-MM-dd}' and tbl_patient.cmp_id=" + cmpid;
            DataTable dt = service.GetData(sql);
            
            string template = Path.Combine(_environment.WebRootPath, "Templates/CCDA.xml");

            XMLZipHelper helper = new XMLZipHelper();

            //byte[] zip = helper.GenerateZip(dt, template);
            using (MemoryStream ms = new MemoryStream())
            {
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string content = System.IO.File.ReadAllText(template);

                        // Replace placeholders
                        foreach (DataColumn col in row.Table.Columns)
                        {
                            string placeholder = $"`{col.ColumnName}`";
                            string value = row[col]?.ToString() ?? "";
                            content = content.Replace(placeholder, value);
                        }

                        string fileName =
                            row["lname"].ToString() + "_" + row["fname"].ToString() + ".xml";

                        var entry = zip.CreateEntry(fileName);

                        using (var stream = entry.Open())
                        using (var writer = new StreamWriter(stream, Encoding.UTF8))
                        {
                            writer.Write(content);
                        }
                    }
                }


                 zipbytes = ms.ToArray();
            }
            return File(
            zipbytes,
            "application/zip",
            $"CCDAXml_{fromDate:yyyyMMdd}_{toDate:yyyyMMdd}.zip");
        }

        public IActionResult DownloadXml(int pId)
        {
            ParentService service = new ParentService();

            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

            string sql = $@"
        SELECT
            tbl_patient.*,
            tbl_patient_ie.*,
            tbl_ie_page1.*,
            DATE_FORMAT(tbl_patient.dob, '%Y%m%d') AS birthtime,
            CASE
                WHEN tbl_patient.gender = 1 THEN 'M'
                WHEN tbl_patient.gender = 2 THEN 'F'
                ELSE 'Other'
            END AS gendercode
        FROM tbl_patient
        INNER JOIN tbl_patient_ie
            ON tbl_patient.id = tbl_patient_ie.patient_id
        LEFT JOIN tbl_ie_page1
            ON tbl_patient.id = tbl_ie_page1.patient_id
        WHERE tbl_patient_ie.patient_id = " + pId ;

            DataTable dt = service.GetData(sql);

            if (dt.Rows.Count == 0)
                return NotFound();

            string template = Path.Combine(_environment.WebRootPath, "Templates/CCDA.xml");

            string content = System.IO.File.ReadAllText(template);

            DataRow row = dt.Rows[0];

            foreach (DataColumn col in dt.Columns)
            {
                string placeholder = $"`{col.ColumnName}`";
                string value = row[col]?.ToString() ?? "";
                content = content.Replace(placeholder, value);
            }

            string fileName = $"{row["lname"]}_{row["fname"]}.xml";

            return File(
                Encoding.UTF8.GetBytes(content),
                "application/xml",
                fileName);
        }
    }
}
