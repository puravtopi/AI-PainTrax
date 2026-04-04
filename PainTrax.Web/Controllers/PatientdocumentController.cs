using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using MS.Services;
using Newtonsoft.Json;
using Optivem.Framework.Core.Domain;
using PainTrax.Services;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Xceed.Words.NET;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;



namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class PatientdocumentController : Controller
    {
        private readonly ILogger<PatientdocumentController> _logger;
        private readonly IMapper _mapper;
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;
        private string _outputPath;
        private string _storagePath;
        private readonly PatientService _patientservices = new PatientService();

        public PatientdocumentController(IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
                                ILogger<PatientdocumentController> logger, IConfiguration configuration)
        {
            _mapper = mapper;
            _logger = logger;
            Environment = environment;
            Configuration = configuration;
            _outputPath = Path.Combine(environment.ContentRootPath, "wwwroot/Content");
            _storagePath = Path.Combine(environment.ContentRootPath, "PatientDocuments");

        }
        public ActionResult Index(int id = 0)
        {
            {
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, id);

                var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments");
                bool folderExists = Directory.Exists(FolderPath);
                if (!folderExists)
                    Directory.CreateDirectory(FolderPath);

                string[] dirs = Directory.GetDirectories(FolderPath, "*", SearchOption.TopDirectoryOnly);

                List<TreeViewNode> nodes = new List<TreeViewNode>();

                int i = 0;
                foreach (var item in dirs)
                {
                    i++;
                    string FolderName = System.IO.Path.GetFileName(item);
                    var FolderPathFile = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", FolderName.ToString(), id.ToString());
                    int j = 0;

                    bool folderExistsNew = Directory.Exists(FolderPathFile);
                    if (!folderExistsNew)
                        Directory.CreateDirectory(FolderPathFile);

                    foreach (var item1 in Directory.GetFiles(FolderPathFile))
                    {
                        j++;
                        nodes.Add(new TreeViewNode { id = j.ToString() + "-" + i.ToString() + "~" + FolderName.ToString() + "$" + System.IO.Path.GetFileName(item1), parent = i.ToString(), text = System.IO.Path.GetFileName(item1) });
                    }

                    nodes.Add(new TreeViewNode { id = i.ToString(), parent = "#", text = FolderName.ToString() + "(" + j + ")" });
                }

                ViewBag.Json = JsonConvert.SerializeObject(nodes, Formatting.Indented);
                return View();
            }

        }
       
        [HttpPost]
        public ActionResult Index(string selectedItems)
        {

            List<TreeViewNode> items = JsonConvert.DeserializeObject<List<TreeViewNode>>(selectedItems);
            return RedirectToAction("Index");
        }
        public async Task<IActionResult> DownloadFilesAsZip(string filenames)
        {
            string PatientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
            var patientData = _patientservices.GetOne(Convert.ToInt32(PatientID));
            string selectedFolder = "";
            if (string.IsNullOrEmpty(filenames))
            {
                return Content("No filenames provided.");
            }

            var fileList = filenames.Split(';').ToList();

            // Create a memory stream to hold the ZIP file in memory
            var memoryStream = new MemoryStream();

            // Create a ZipArchive and use it to add files to the memory stream
            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, leaveOpen: true)) // Leave stream open
            {
                foreach (var filename in fileList)
                {
                    // Extract details about the file (you can adjust the parsing logic as needed)
                    string patientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId)?.ToString();
                    string split = filename.Split('~')[0];
                    string split1 = filename.Split('~')[1];
                    selectedFolder = split1.Split('$')[0];
                    string fileNameNew = split1.Split('$')[1];

                    var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", selectedFolder, patientID);
                    var filePath = Path.Combine(folderPath, fileNameNew);

                    if (System.IO.File.Exists(filePath))
                    {
                        // Add the file to the ZIP archive
                        var zipEntry = archive.CreateEntry(fileNameNew, CompressionLevel.Fastest);
                        using (var entryStream = zipEntry.Open())
                        using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        {
                            await fileStream.CopyToAsync(entryStream);
                        }
                    }
                    else
                    {
                        // Return an error if any file does not exist
                        return Content($"File not found: {filename}");
                    }
                }
            }

            // Set the position of the memory stream back to the beginning before returning it
            memoryStream.Position = 0;

            // Return the ZIP file as a FileResult
            return File(memoryStream, "application/zip", patientData.fname + "_" +patientData.lname + ".zip");
        }
        

       
        

        public async Task<IActionResult> Download(string filename)
        {
            string PatientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
            string selectedfolder = filename.Split('~')[0];
            string filenamenew = filename.Split('~')[1];
            if (filenamenew == null)
                return Content("filename is not availble");
            var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", selectedfolder, PatientID);

            var path = Path.Combine(FolderPath, filenamenew);

            var memory = new MemoryStream();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            return File(memory, GetContentType(path), Path.GetFileName(path));
        }

        public async Task<IActionResult> DeleteDoc(string filename)
        {
            string PatientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
            string selectedfolder = filename.Split('~')[0];
            string filenamenew = filename.Split('~')[1];
            if (filenamenew == null)
                return Content("filename is not availble");
            var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", selectedfolder, PatientID);
            var path = Path.Combine(FolderPath, filenamenew);

            if (path != null || path != string.Empty)
            {
                if ((System.IO.File.Exists(path)))
                {
                    System.IO.File.Delete(path);
                }

            }
            //return RedirectToAction("Index/" + PatientID + "", "Patientdocument");
            return RedirectToAction("Index", "Patientdocument", new { id = PatientID });
            // return RedirectToAction("Index");
        }
        public IActionResult PreviewDoc(string filename)
        {
            string PatientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
            string selectedfolder = filename.Split('~')[0];
            string filenamenew = filename.Split('~')[1];

            if (string.IsNullOrEmpty(filenamenew))
                return Content("Filename is not available");

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", selectedfolder, PatientID);
            var path = Path.Combine(folderPath, filenamenew);

            if (!System.IO.File.Exists(path))
                return NotFound();

            var contentType = GetContentType(path);

            ViewBag.ContentType = contentType;
            ViewBag.Filename = filename;

            if (contentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                // Extract content from the Word document and generate HTML
                string htmlContent = ConvertWordToHtml(path);
                ViewBag.DocContent = htmlContent;
            }

            return View();
        }
        public IActionResult GetFile(string filename)
        {
            string PatientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
            string selectedfolder = filename.Split('~')[0];
            string filenamenew = filename.Split('~')[1];

            if (string.IsNullOrEmpty(filenamenew))
                return Content("Filename is not available");

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", selectedfolder, PatientID);
            var path = Path.Combine(folderPath, filenamenew);

            if (!System.IO.File.Exists(path))
                return NotFound();

            var contentType = GetContentType(path);

            var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            return File(fileStream, contentType);
        }

        [HttpPost]
        public IActionResult ImportData(string selectedParent, IFormFile[] postedFile)
        {
            string PatientID = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
            try
            {
                string SelectedFolder = selectedParent.Replace('"', ' ').Trim();
                if (postedFile != null && postedFile.Length > 0)
                {
                    foreach (var file in postedFile)
                    {
                        // Get the file name from the browser
                        var fileName = System.IO.Path.GetFileName(file.FileName);


                        // Get file path to be uploaded

                        var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", SelectedFolder, PatientID);


                        bool folderExists = Directory.Exists(FolderPath);
                        if (!folderExists)
                            Directory.CreateDirectory(FolderPath);

                        var filePath = Path.Combine(FolderPath, fileName);

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

                    }
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "List");
            }
            return RedirectToAction("Index", "Patientdocument", new { id = PatientID });
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        } //end GetColumnName method

        /// <summary>
        /// Given just the column name (no row index), it will return the zero based column index.
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
        /// A length of three can be implemented when needed.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
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

        #region private Method
        private void SaveLog(Exception ex, string acctionname)
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
                Message = msg,
            };
            new LogService().Insert(logdata);
        }

        private string GetContentType(string path)
        {
            var provider = new FileExtensionContentTypeProvider();
            if (!provider.TryGetContentType(path, out var contentType))
            {
                contentType = "application/octet-stream";
            }
            return contentType;
        }

        //private string GetContentType(string path)
        //{
        //    var types = GetMimeTypes();
        //    var ext = Path.GetExtension(path).ToLowerInvariant();
        //    return types[ext];
        //}
        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
            {
                {".txt", "text/plain"},
                {".pdf", "application/pdf"},
                {".doc", "application/vnd.ms-word"},
                {".docx", "application/vnd.ms-word"},
                {".xls", "application/vnd.ms-excel"},
                {".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
                {".jpeg", "image/jpeg"},
                {".gif", "image/gif"},
                {".csv", "text/csv"}
            };
        }
        private string ConvertWordToHtml(string filePath)
        {
            StringBuilder sb = new StringBuilder();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                Body body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var element in body.Elements())
                {
                    sb.Append(GetHtmlFromElement(element));
                }
            }

            // Wrap the generated HTML content with a <div> and apply some basic styles
            return $"<div style=\"font-family: Arial; font-size: 14px;\">{sb}</div>";
        }

        private string GetHtmlFromElement(OpenXmlElement element)
        {
            StringBuilder sb = new StringBuilder();

            if (element is Paragraph)
            {
                sb.Append("<p>");

                foreach (var run in element.Elements<Run>())
                {
                    sb.Append(GetHtmlFromRun(run));
                }

                sb.Append("</p>");
            }
            else if (element is Run)
            {
                sb.Append(GetHtmlFromRun(element));
            }
            // Add handling for other element types as needed

            return sb.ToString();
        }

        private string GetHtmlFromRun(OpenXmlElement run)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<span>");

            // Extract text content from the run element
            foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
            {
                sb.Append(text.Text);
            }

            sb.Append("</span>");

            return sb.ToString();
        }

        /* PatientDocumentController.cs replace oldpath = @"D:\Apex\LiveServerOld\AKS - Copy - AKS - Copy\PatientDocument" with server path  */
        public ActionResult Transfer()
        {
            //string oldpath = @"D:\Apex\LiveServerOld\AKS - Copy - AKS - Copy\PatientDocument";
            string oldpath = @"E:\ProductionServer\ePainTrax_NV_AKS\PatientDocument";
            string[] dirs = Directory.GetDirectories(oldpath, "*", SearchOption.TopDirectoryOnly);
            List<Tuple<string, string, string>> folderList = new List<Tuple<string, string, string>>();
            ParentService _service = new ParentService();
            foreach (string folder in dirs)
            {
                string foldername = System.IO.Path.GetFileName(folder);
                string newfoldername = "";
                string status = "Transfered";
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                //folderList.Add(FolderName);
                DataTable dt = _service.GetData("select id from tbl_patient where old_id=" + foldername + " and cmp_id=" + cmpid);
                if (dt.Rows.Count > 0)
                {
                    newfoldername = dt.Rows[0]["id"].ToString();
                    status = !Directory.Exists(Path.Combine(_storagePath + "/Old Documents/" + newfoldername)) ? "Directory Not Created" : status;
                    if (Directory.Exists(Path.Combine(_storagePath + "/Old Documents/" + newfoldername)))
                    {
                        string[] sourceFiles = Directory.GetFiles(Path.Combine(oldpath, foldername));
                        foreach (string sourceFile in sourceFiles)
                        {
                            string fileName = Path.GetFileName(sourceFile);
                            string targetFile = Path.Combine(_storagePath + "/Old Documents/" + newfoldername, fileName);

                            if (!System.IO.File.Exists(targetFile))
                            {
                                status = "Partial Copied";
                            }
                        }
                    }
                }
                else
                {
                    status = "Old Id Not Found";
                }

                folderList.Add(Tuple.Create(foldername, newfoldername, status));
            }
            ViewBag.Folders = folderList;
            return View();
        }

        [HttpPost]
        public JsonResult TransferProcess(string str)
        {
            string oldpath = @"E:\ProductionServer\ePainTrax_NV_AKS\PatientDocument";
            string[] dirs = Directory.GetDirectories(oldpath, "*", SearchOption.TopDirectoryOnly);
            ParentService _service = new ParentService();
            foreach (string folder in dirs)
            {
                string foldername = System.IO.Path.GetFileName(folder);
                string newfoldername = "";
                string status = "";
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                DataTable dt = _service.GetData("select id from tbl_patient where old_id=" + foldername + " and  cmp_id=" + cmpid);
                if (dt.Rows.Count > 0)
                {
                    newfoldername = dt.Rows[0]["id"].ToString();
                    if (!Directory.Exists(Path.Combine(_storagePath + "/Old Documents/" + newfoldername)))
                    {
                        Directory.CreateDirectory(Path.Combine(_storagePath + "/Old Documents/" + newfoldername));
                    }
                    string[] sourceFiles = Directory.GetFiles(Path.Combine(oldpath, foldername));
                    foreach (string sourceFile in sourceFiles)
                    {
                        string fileName = Path.GetFileName(sourceFile);
                        string targetFile = Path.Combine(_storagePath + "/Old Documents/" + newfoldername, fileName);

                        if (!System.IO.File.Exists(targetFile))
                        {
                            System.IO.File.Copy(sourceFile, targetFile);
                        }
                    }

                }
                else
                {
                    status = "Old Id Not Found";
                }


            }

            return Json(new { status = 1 });

        }


        #endregion
    }
}
