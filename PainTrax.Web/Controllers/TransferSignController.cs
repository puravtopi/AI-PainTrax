using Microsoft.AspNetCore.Mvc;
using PainTrax.Services;
using PainTrax.Web.Helper;
using System.Data;

namespace PainTrax.Web.Controllers
{
    public class TransferSignController : Controller
    {
        private readonly ILogger<PatientdocumentController> _logger;
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment _environment;
        private IConfiguration _configuration;
        private string _storagePath;

        private readonly ParentService _pareentservices = new ParentService();

        public TransferSignController(Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
                                ILogger<PatientdocumentController> logger, IConfiguration configuration)
        {

            _logger = logger;
            _environment = environment;
            _configuration = configuration;
            _storagePath = Path.Combine(environment.WebRootPath, "signatures");

        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult TransferBHF()
        {

            string filePath = Path.Combine(_environment.WebRootPath, "Logfiles/TransferSign.txt");

            // Create writer (false = overwrite, true = append)
            StreamWriter writer = new StreamWriter(filePath, false);
        
            //string _sourcePath = @"D:\Apex_Online\ePainTrax_NV_BHF_V\Sign";
            string _sourcePath = @"E:\ProductionServer\ePainTrax_NV_BHF_V\Sign";
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            DataTable dt = new DataTable();
            List<string> Patients = new List<string>();
            dt = _pareentservices.GetData($"select * from tbl_patient where  cmp_id={cmpid}");
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    string old_id = row["old_id"].ToString();
                    string new_id = row["id"].ToString();
                    string name = row["lname"].ToString() + "," + row["fname"].ToString();
                    writer.WriteLine($" Old Id:{old_id} , New Id:{new_id}, Name:{name} ");
                    //string folderPath = _sourcePath + "/" + old_id;

                    if (!string.IsNullOrEmpty(old_id))
                    {
                        string? oldfile = Directory.GetFiles(_sourcePath, $"{old_id}_*.jp*g").FirstOrDefault();

                        if (oldfile != null)
                        {
                            writer.WriteLine($"Old file: {oldfile} ");
                            string ext = Path.GetExtension(oldfile);   // .jpg or .jpeg
                            string newFile = Path.Combine(_storagePath, new_id + ext);
                            try
                            {
                                System.IO.File.Copy(oldfile, newFile);
                                    DateTime newDateTime = DateTime.Now;   // or any date

                                System.IO.File.SetCreationTime(newFile, newDateTime);
                                System.IO.File.SetLastWriteTime(newFile, newDateTime);
                                System.IO.File.SetLastAccessTime(newFile, newDateTime);
                                writer.WriteLine($"Copy  Old File:{oldfile} to New File:{newFile}");
                            }
                            catch (Exception ex)
                            {
                                writer.WriteLine($" Error :" + ex.Message);
                            }
                            
                        }
                        else
                        {
                            writer.WriteLine($" Error :" +  old_id + "  Sign not found ");
                        }
                    }
                }
            }
            TempData["message"] = "Complate Tranfering Signature";
            TempData["alert"] = "alert alert-success";
            writer.Close();
            return RedirectToAction("Index");

        }
    }

}

