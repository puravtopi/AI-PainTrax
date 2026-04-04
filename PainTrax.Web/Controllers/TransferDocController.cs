using GroupDocs.Viewer.Options;
using Microsoft.AspNetCore.Mvc;
using MS.Services;
using PainTrax.Services;
using PainTrax.Web.Helper;
using System.Data;

namespace PainTrax.Web.Controllers
{
    public class TransferDocController : Controller
    {
        private readonly ILogger<PatientdocumentController> _logger;
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment _environment;
        private IConfiguration _configuration;
        private string _storagePath;
        
        private readonly ParentService _pareentservices = new ParentService();

        public TransferDocController( Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
                                ILogger<PatientdocumentController> logger, IConfiguration configuration)
        {
            
            _logger = logger;
            _environment = environment;
            _configuration = configuration;
            _storagePath = Path.Combine(environment.ContentRootPath, "PatientDocuments/Old Documents");

        }
        public IActionResult Index()
        {
          
            return View();
        }

        [HttpPost]
        public IActionResult TransferSaldesMD()
        {
            
            string filePath = Path.Combine(_environment.WebRootPath, "Logfiles/TransferDoc.txt");

            // Create writer (false = overwrite, true = append)
            StreamWriter writer = new StreamWriter(filePath, false);

            string _sourcePath = @"E:\ProductionServer\ePainTrax_NV_AKS\PatientDocument";
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            DataTable dt = new DataTable();
            List<string>Patients = new List<string>();
            dt = _pareentservices.GetData($"select * from tbl_patient where  cmp_id={cmpid}");
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    string old_id = row["old_id"].ToString();
                    string new_id = row["id"].ToString();
                    string name = row["lname"].ToString()+","+ row["fname"].ToString();
                    writer.WriteLine($" Old Id:{old_id} , New Id:{new_id}, Name:{name} ");
                    string folderPath = _sourcePath+"/"+old_id;

                    if (Directory.Exists(folderPath))
                    {
                        string[] files = Directory.GetFiles(folderPath);

                        foreach (string file in files)
                        {
                            string fileName = Path.GetFileName(file);  // only name, no full path
                            writer.WriteLine($" {old_id}: {fileName}");
                            if(!Directory.Exists(_storagePath+"/"+new_id))
                            {
                                writer.WriteLine("Directory Created : "+_storagePath + "/" + new_id);
                                Directory.CreateDirectory(_storagePath+"/"+new_id);
                            }
                            if(System.IO.File.Exists(_storagePath + "/" + new_id+"/"+fileName))
                            {
                                writer.WriteLine("File already Exsits : " + _storagePath + "/" + new_id + "/" + fileName);
                            }
                            else
                            {
                                try
                                {
                                    writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName);
                                    System.IO.File.Copy(file, _storagePath + "/" + new_id + "/" + fileName);
                                }catch (Exception ex) { writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName+" Error :" + ex.ToString()); }
                            }

                        }
                    }
                }
            }
            TempData["message"] = "Complate Tranfering Data";
            TempData["alert"] = "alert alert-success";
            writer.Close();
            return RedirectToAction("Index");
         
        }

        //[HttpPost]
        //public IActionResult TransferBHFTest()
        //{

        //    string filePath = Path.Combine(_environment.WebRootPath, "Logfiles/TransferDocBHF.txt");

        //    // Create writer (false = overwrite, true = append)
        //    StreamWriter writer = new StreamWriter(filePath, false);

        //    string _sourcePath = @"E:\ProductionServer\ePainTrax_NV_BHF_V\PatientDocument";
        //    string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
        //    DataTable dt = new DataTable();
        //    List<string> Patients = new List<string>();
        //    dt = _pareentservices.GetData($"select * from tbl_patient where  cmp_id={cmpid}");
        //    if (dt.Rows.Count > 0)
        //    {
        //        foreach (DataRow row in dt.Rows)
        //        {
        //            string old_id = row["old_id"].ToString();
        //            string new_id = row["id"].ToString();
        //            string name = row["lname"].ToString() + "," + row["fname"].ToString();
        //            writer.WriteLine($" Old Id:{old_id} , New Id:{new_id}, Name:{name} ");
        //            string folderPath = _sourcePath + "/" + old_id;

        //            if (Directory.Exists(folderPath))
        //            {
        //                string[] files = Directory.GetFiles(folderPath);

        //                foreach (string file in files)
        //                {
        //                    string fileName = Path.GetFileName(file);  // only name, no full path
        //                    writer.WriteLine($" {old_id}: {fileName}");
        //                    if (!Directory.Exists(_storagePath + "/" + new_id))
        //                    {
        //                        writer.WriteLine("Directory Created : " + _storagePath + "/" + new_id);
        //                        Directory.CreateDirectory(_storagePath + "/" + new_id);
        //                    }
        //                    if (System.IO.File.Exists(_storagePath + "/" + new_id + "/" + fileName))
        //                    {
        //                        writer.WriteLine("File already Exsits : " + _storagePath + "/" + new_id + "/" + fileName);
        //                    }
        //                    else
        //                    {
        //                        try
        //                        {
        //                            writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName);
        //                            System.IO.File.Copy(file, _storagePath + "/" + new_id + "/" + fileName);
        //                        }
        //                        catch (Exception ex) { writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName + " Error :" + ex.ToString()); }
        //                    }

        //                }
        //            }
        //        }
        //    }
        //    TempData["message"] = "Complate Tranfering Data";
        //    TempData["alert"] = "alert alert-success";
        //    writer.Close();
        //    return RedirectToAction("Index");

        //}

        [HttpPost]
        public IActionResult TransferBHFPC()
        {

            string filePath = Path.Combine(_environment.WebRootPath, "Logfiles/TransferDocBHF.txt");

            // Create writer (false = overwrite, true = append)
            StreamWriter writer = new StreamWriter(filePath, false);

            string _sourcePath = @"E:\ProductionServer\ePainTrax_NV_BHF_V\PatientDocument";
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
                    string folderPath = _sourcePath + "/" + old_id;

                    if (Directory.Exists(folderPath))
                    {
                        string[] files = Directory.GetFiles(folderPath);

                        foreach (string file in files)
                        {
                            string fileName = Path.GetFileName(file);  // only name, no full path
                            writer.WriteLine($" {old_id}: {fileName}");
                            if (!Directory.Exists(_storagePath + "/" + new_id))
                            {
                                writer.WriteLine("Directory Created : " + _storagePath + "/" + new_id);
                                Directory.CreateDirectory(_storagePath + "/" + new_id);
                            }
                            if (System.IO.File.Exists(_storagePath + "/" + new_id + "/" + fileName))
                            {
                                writer.WriteLine("File already Exsits : " + _storagePath + "/" + new_id + "/" + fileName);
                            }
                            else
                            {
                                try
                                {
                                    writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName);
                                    System.IO.File.Copy(file, _storagePath + "/" + new_id + "/" + fileName);
                                }
                                catch (Exception ex) { writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName + " Error :" + ex.ToString()); }
                            }

                        }
                    }
                }
            }
            TempData["message"] = "Complate Tranfering Data";
            TempData["alert"] = "alert alert-success";
            writer.Close();
            return RedirectToAction("Index");

        }

        [HttpPost]
        public IActionResult TransferQMPPC()
        {

            string filePath = Path.Combine(_environment.WebRootPath, "Logfiles/TransferQMPPCDoc.txt");

            // Create writer (false = overwrite, true = append)
            StreamWriter writer = new StreamWriter(filePath, false);

            string _sourcePath = @"E:\ProductionServer\ePainTrax_NV_QMPPC\PatientDocument";
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
                    string folderPath = _sourcePath + "/" + old_id;

                    if (Directory.Exists(folderPath))
                    {
                        string[] files = Directory.GetFiles(folderPath);

                        foreach (string file in files)
                        {
                            string fileName = Path.GetFileName(file);  // only name, no full path
                            writer.WriteLine($" {old_id}: {fileName}");
                            if (!Directory.Exists(_storagePath + "/" + new_id))
                            {
                                writer.WriteLine("Directory Created : " + _storagePath + "/" + new_id);
                                Directory.CreateDirectory(_storagePath + "/" + new_id);
                            }
                            if (System.IO.File.Exists(_storagePath + "/" + new_id + "/" + fileName))
                            {
                                writer.WriteLine("File already Exsits : " + _storagePath + "/" + new_id + "/" + fileName);
                            }
                            else
                            {
                                try
                                {
                                    writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName);
                                    System.IO.File.Copy(file, _storagePath + "/" + new_id + "/" + fileName);
                                }
                                catch (Exception ex) { writer.WriteLine("File Copy : " + file + " To " + _storagePath + "/" + new_id + "/" + fileName + " Error :" + ex.ToString()); }
                            }

                        }
                    }
                }
            }
            TempData["message"] = "Complate Tranfering Data";
            TempData["alert"] = "alert alert-success";
            writer.Close();
            return RedirectToAction("Index");

        }
    }
}
