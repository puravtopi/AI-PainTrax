using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using MS.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System;
//using iText.Forms.Fields;
//using iText.Kernel.Pdf;
//using iText.Forms;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class FormsController : Controller
    {
        private readonly IMapper _mapper;
        private readonly ILogger<FormsController> _logger;

        private readonly IWebHostEnvironment _environment;
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly PatientService _patientservices = new PatientService();
       

        public FormsController(ILogger<FormsController> logger, IWebHostEnvironment environment)
        {
            _environment = environment;
            _logger = logger;
        }
        public IActionResult Index(string searchtxt = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();

            string cnd = " and cmp_id=" + cmpid;

            if (!string.IsNullOrEmpty(searchtxt))
                cnd = " and title like '%" + searchtxt + "%' ";

            var result = _patientservices.GetAll(cnd);
            var data = result;

            var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid.TrimEnd());
            if (!Directory.Exists(downloadFolder))
            {
                Directory.CreateDirectory(downloadFolder);
                Directory.CreateDirectory(downloadFolder + "/Forms");
                Directory.CreateDirectory(downloadFolder + "/BS");
                Directory.CreateDirectory(downloadFolder + "/Letters");
                Directory.CreateDirectory(downloadFolder + "/Others");
            }
            else
            {
                if (!Directory.Exists(downloadFolder + "/Forms")) Directory.CreateDirectory(downloadFolder + "/Forms");
                if (!Directory.Exists(downloadFolder + "/BS")) Directory.CreateDirectory(downloadFolder + "/BS");
                if (!Directory.Exists(downloadFolder + "/Letters")) Directory.CreateDirectory(downloadFolder + "/Letters");
                if (!Directory.Exists(downloadFolder + "/Others")) Directory.CreateDirectory(downloadFolder + "/Others");

            }

            //string[] formsPaths = Directory.GetFiles(downloadFolder + "/Forms");
            //string[] bsPaths = Directory.GetFiles(downloadFolder + "/BS");
            //string[] lettersPaths = Directory.GetFiles(downloadFolder + "/Letters");
            //string[] othersPaths = Directory.GetFiles(downloadFolder + "/Others");

            //List<string> pdffiles = new List<string>();
            //List<string> bsfiles = new List<string>();
            //List<string> lettersfiles = new List<string>();
            //List<string> othersfiles = new List<string>();
            //for (int i = 0; i < formsPaths.Length; i++)
            //{
            //    if (formsPaths[i].ToLower().EndsWith("pdf"))
            //        pdffiles.Add(formsPaths[i].Substring(formsPaths[i].LastIndexOf("\\") + 1));
            //}
            //for (int i = 0; i < bsPaths.Length; i++)
            //{
            //    if (bsPaths[i].ToLower().EndsWith("pdf"))
            //        bsfiles.Add(bsPaths[i].Substring(bsPaths[i].LastIndexOf("\\") + 1));
            //}
            //for (int i = 0; i < lettersPaths.Length; i++)
            //{
            //    if (lettersPaths[i].ToLower().EndsWith("pdf"))
            //        lettersfiles.Add(lettersPaths[i].Substring(lettersPaths[i].LastIndexOf("\\") + 1));
            //}
            //for (int i = 0; i < othersPaths.Length; i++)
            //{
            //    if (othersPaths[i].ToLower().EndsWith("pdf"))
            //        othersfiles.Add(othersPaths[i].Substring(othersPaths[i].LastIndexOf("\\") + 1));
            //}
            //ViewBag.pdffiles = pdffiles;
            //ViewBag.bsfiles = bsfiles;
            //ViewBag.lettersfiles = lettersfiles;
            //ViewBag.othersfiles = othersfiles;

            var subFolders = Directory.GetDirectories(downloadFolder);

            var filesByFolder = new Dictionary<string, List<string>>();

            var formAccess = HttpContext.Session.GetString(SessionKeys.SessionFormsAccess);

            List<string> fileList=new List<string>();

            foreach (var folder in subFolders)
            {
                fileList = new List<string>();
                var folderName = Path.GetFileName(folder);
                var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                        .Select(Path.GetFileName)
                                        .ToList();

                foreach (var item in pdfFiles)
                {
                    if (formAccess.Contains(item))
                    {
                        fileList.Add(item);
                    }
                }
                filesByFolder.Add(folderName, fileList);
            }

            ViewBag.FilesByFolder = filesByFolder;


            return View(data);

        }
        public IActionResult Manage(string searchtxt = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();

            string cnd = " and cmp_id=" + cmpid;

            if (!string.IsNullOrEmpty(searchtxt))
                cnd = " and title like '%" + searchtxt + "%' ";

            var result = _patientservices.GetAll(cnd);
            var data = result;

            var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
            var subFolders = Directory.GetDirectories(downloadFolder);

            var filesByFolder = new Dictionary<string, List<string>>();

            foreach (var folder in subFolders)
            {
                var folderName = Path.GetFileName(folder);
                var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                        .Select(Path.GetFileName)
                                        .ToList();

                filesByFolder.Add(folderName, pdfFiles);
            }

            ViewBag.FilesByFolder = filesByFolder;
            ViewBag.FolderNames = filesByFolder.Keys.ToList();

            return View();

        }

        [HttpPost]
        public ActionResult CreateFolder(string folderName)
        {

            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
            string newFolderPath = Path.Combine(downloadFolder, folderName);
            if (!Directory.Exists(newFolderPath))
            {
                Directory.CreateDirectory(newFolderPath);
                TempData["Message"] = "Folder created successfully.";
            }
            else
            {
                TempData["Message"] = "Folder already exists.";
            }

            return RedirectToAction("Manage");
        }

        [HttpPost]
        public ActionResult DeleteFolder(string folderName)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
            string folderPath = Path.Combine(downloadFolder, folderName);
            if (Directory.Exists(folderPath))
            {
                if (!Directory.EnumerateFileSystemEntries(folderPath).Any())
                {
                    Directory.Delete(folderPath);
                    TempData["Message"] = "Folder deleted successfully.";
                }
                else
                {
                    TempData["Message"] = "Folder is not empty.";
                }
            }
            else
            {
                TempData["Message"] = "Folder does not exist.";
            }

            return RedirectToAction("Manage");
        }

        [HttpPost]
        public ActionResult UploadFile(string folderName, IFormFile file)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            var downloadFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
            if (file != null && file.Length > 0 && Path.GetExtension(file.FileName).ToLower() == ".pdf")
            {
                string folderPath = Path.Combine(downloadFolder, folderName);
                if (Directory.Exists(folderPath))
                {
                    string filePath = Path.Combine(folderPath, Path.GetFileName(file.FileName));
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                    TempData["Message"] = "File uploaded successfully.";
                }
                else
                {
                    TempData["Message"] = "Selected folder does not exist.";
                }
            }
            return RedirectToAction("Manage");

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
        public IActionResult GeneratePdf(string pdffile, string id, string txt_date = "", string txt_surgery = "", string txt_docName = "", string txt_MCode_Proc = "", string txtProcedureCode = "", string txt_diagnosis = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            Dictionary<string, string> controls = new Dictionary<string, string>();

            try
            {
                if (txt_date != null && txt_date.Trim() != "") controls.Add("txt_date", txt_date);
                if (txt_surgery != null && txt_surgery.Trim() != "") controls.Add("txt_surgery", txt_surgery);
                if (txt_docName != null && txt_docName.Trim() != "") controls.Add("txt_docName", txt_docName);
                if (txt_MCode_Proc != null && txt_MCode_Proc.Trim() != "") controls.Add("txt_MCode_Proc", txt_MCode_Proc);
                if (txtProcedureCode != null && txtProcedureCode.Trim() != "") controls.Add("txtProcedureCode", txtProcedureCode);
                if (txt_diagnosis != null && txt_diagnosis.Trim() != "") controls.Add("txt_diagnosis", txt_diagnosis);

            }
            catch (Exception ex)
            {
                SaveLog(ex, "Read Textboxes");
            }

            PdfHelper _pdfhelper = new PdfHelper();
            string outputfilename = "";
            var uploadsFolder = "";
            var filePath = "";
            var signPath = "";

            try
            {
                uploadsFolder = Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid);
                filePath = Path.Combine(uploadsFolder, pdffile);

                signPath = Path.Combine(_environment.WebRootPath, "signatures");
            }
            catch (Exception ex)
            {
                SaveLog(ex, "set Paths");
            }

            byte[] pdfBytes = null;
            //try
            //{
            //    // Create a memory stream to hold the modified PDF
            //    using (MemoryStream outputStream = new MemoryStream())
            //    {
            //        // Create a PdfReader object to read the input PDF
            //        using (PdfReader reader = new PdfReader(filePath))
            //        {
            //            // Create a PdfWriter object to write to the output PDF
            //            using (PdfWriter writer = new PdfWriter(outputStream))
            //            {
            //                // Create a PdfDocument object
            //                using (PdfDocument pdfDoc = new PdfDocument(reader, writer))
            //                {
            //                    // Get the form fields
            //                    PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);

            //                    // Get a specific text field by its name
            //                    PdfFormField field = form.GetField("#fname");

            //                    if (field != null)
            //                    {
            //                        // Set the field value
            //                        field.SetValue("Test Value");
            //                    }
            //                    else
            //                    {
            //                        Console.WriteLine("Field not found.");
            //                    }
            //                }
            //            }
            //        }

            //        // Get the byte array of the modified PDF from the memory stream
            //        pdfBytes = outputStream.ToArray();

            //        // At this point, you can do whatever you want with the modified PDF bytes.
            //        // For example, you can save them to a file, send them over the network, etc.
            //        // In this example, we're just printing out the length of the modified PDF bytes.
            //       // Console.WriteLine("Modified PDF size: " + pdfBytes.Length + " bytes");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Error: " + ex.Message);
            //}




            try
            {
                pdfBytes = _pdfhelper.Stamping(filePath, "Id", id, controls, cmpid, signPath);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Pdf Stamping");
            }

            //string htmlContent = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "demo.html"));

            return File(pdfBytes, "application/pdf");
        }

        public IActionResult DisplayColumns()
        {

            var viewStructure = _patientservices.GetData("DESCRIBE View_Pdf;");
            ViewBag.ViewStructure = viewStructure;
            return View();
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
