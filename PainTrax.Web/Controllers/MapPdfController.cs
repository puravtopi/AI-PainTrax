using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Server.HttpSys;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using MS.Services;
using PainTrax.Web.Helper;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Xml.Linq;
//using static iTextSharp.awt.geom.Point2D;

namespace PainTrax.Web.Controllers
{
    public class MapPdfController : Controller
    {
        private readonly IWebHostEnvironment _environment;
        private readonly PatientService _patientservices = new PatientService();
        private readonly AppHelper _apphelper = new AppHelper();
        public MapPdfController(IWebHostEnvironment environment)
        {
            _environment = environment;
        }
        public IActionResult Index()
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf/" + cmpid);
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            string[] filePaths = Directory.GetFiles(uploadsFolder);
            List<string> pdffiles=new List<string>();
        
            for (int i = 0; i < filePaths.Length; i++)
            {
                if (filePaths[i].ToLower().EndsWith("pdf"))
                    pdffiles.Add(filePaths[i].Substring(filePaths[i].LastIndexOf("\\") + 1));
            }
            ViewBag.pdffiles = pdffiles;
            DataTable dt = _patientservices.GetData("show tables");
            List<string> tables = new List<string>();
            for (int i=0;i<dt.Rows.Count;i++)
            {
                tables.Add(dt.Rows[i][0].ToString());
            }
            ViewBag.tables = tables;
            return View();
        }

        [HttpPost]
        public IActionResult UploadFile(IFormFile file )
        {
            
            if (file == null || file.Length == 0)
            {
                return BadRequest("File not selected or empty.");
            }
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf/" + cmpid);

            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            var fileName = Path.GetFileName(file.FileName);
            var filePath = Path.Combine(uploadsFolder, fileName);
            
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                 file.CopyTo(stream);
            }

            return Ok($"File '{fileName}' uploaded successfully.");
            
        }

        [HttpPost]
        public IActionResult SelectPdf() 
        {
            return View();
        }

        
        public string SelectTable(string tablename)
        {

            //DataTable dt = _patientservices.GetData("show COLUMNS from tbl_patient;");
            try
            {


                if (tablename != "-- Select --")
                {
                    string data = _apphelper.GetJson("show COLUMNS from " + tablename);

                    return data;
                }
                else
                {
                    return "[]";
                }
            }
            catch (Exception ex)
            {
                return "[]";
            }
        }

        
        public string SelectText(string  pdfname)
        {
            try
            {

                DataTable dt = new DataTable();
                dt.Columns.Add("Textbox", typeof(string));

                if (pdfname != "-- Select --")
                {
                    string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

                    var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf/"+ cmpid);
                    var filePath = Path.Combine(uploadsFolder, pdfname);
                    PdfReader pdfReader = new PdfReader(filePath);
                    foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
                    {
                        if (de.Key.ToString().ToLower() != "txttable" && de.Key.ToString().ToLower() != "txtfile" && de.Key.ToString().ToLower() != "imgsign" && !de.Key.ToString().ToLower().StartsWith("txtfix") && !de.Key.ToString().ToLower().StartsWith("@") && !de.Key.ToString().ToLower().StartsWith("#"))
                        {
                            var row = dt.NewRow();
                            row["Textbox"] =de.Key.ToString() ;
                            dt.Rows.Add(row);
                        }
                    }
                    pdfReader.Close();
                    
                    string json = AppHelper.DataTableToJsonObj(dt);
                    return AppHelper.DataTableToJsonObj(dt);
                }                
                else 
                {
                    return "[]";
                }
            }
            catch (Exception ex)
            {
                return "[]";
            }
        }

        public string SelectMap(string pdfname)
        {
            try
            {

                DataTable dt = new DataTable();
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Value", typeof(string));

                if (pdfname != "-- Select --")
                {
                    string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

                    var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf/" +cmpid);
                    var filePath = Path.Combine(uploadsFolder, pdfname);
                    PdfReader pdfReader = new PdfReader(filePath);
                   // PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(TargetFile, FileMode.Create));
                    AcroFields pdfFormFields = pdfReader.AcroFields;
                    foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
                    {
                        if (de.Key.ToString().ToLower() != "imgsign" && !de.Key.ToString().ToLower().StartsWith("txtfix") && !de.Key.ToString().ToLower().StartsWith("@") && !de.Key.ToString().ToLower().StartsWith("#"))
                        {
                            var row = dt.NewRow();
                            row["Name"] = de.Key.ToString();
                            row["Value"] = pdfFormFields.GetField(de.Key.ToString());
                            dt.Rows.Add(row);
                        }
                    }
                    pdfReader.Close();
                    return AppHelper.DataTableToJsonObj(dt);
                }
                else
                {
                    return "[]";
                }
            }
            catch (Exception ex)
            {
                return "[]";
            }
        }


        [HttpPost]
        public IActionResult SelectField()
        {
            return View();
        }

        
        public string Map(string SourceFile, string TextName="", string ColumnName="",string Table="",string FileName="")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

            var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf/"+cmpid);
            var filePath = Path.Combine(uploadsFolder, SourceFile);
            var TargetFile = Path.Combine(uploadsFolder, "_temp.pdf");
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(filePath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(TargetFile, FileMode.Create));
                AcroFields pdfFormFields = pdfStamper.AcroFields;
                foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
                {
                    if (TextName == "")
                    {
                        if (Table != "")
                        {
                            if (de.Key.ToString().ToLower() == "txttable")
                            {
                                pdfFormFields.SetField(de.Key.ToString(), Table);
                                break;
                            }
                        }
                        if (FileName != "")
                        {
                            if (de.Key.ToString().ToLower() == "txtfile")
                            {
                                pdfFormFields.SetField(de.Key.ToString(), FileName);
                               // break;
                            }
                        }
                    }
                    else
                    {
                        if (ColumnName == "")
                        {
                            if (de.Key.ToString() == TextName)
                            {
                                pdfFormFields.SetField(de.Key.ToString(), "");
                            }
                        }
                        else if (de.Key.ToString() == TextName)
                        {
                            pdfFormFields.SetField(de.Key.ToString(), de.Key.ToString() + "|" + ColumnName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "{\"message\":\"Error\" "+ex.Message+"}";
            }
            finally
            {

                pdfStamper.Close();
                pdfReader.Close();
                System.IO.File.Copy(TargetFile, filePath,true);
            }

            return "{\"message\":\"Success\"}";
        }

       
       /* public string Unmap(string SourceFile, string TextName)
        {
            var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf");
            var filePath = Path.Combine(uploadsFolder, SourceFile);
            //  var SrcFile= Path.Combine(uploadsFolder, "_tempsrc.pdf");
            // System.IO.File.Delete(SrcFile);
            //  System.IO.File.Copy(filePath, SrcFile, true);   
            var TargetFile = Path.Combine(uploadsFolder, "_temp.pdf");

            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {

                pdfReader = new PdfReader(filePath);
                //MemoryStream pdfOutput = new MemoryStream();
                //pdfStamper = new PdfStamper(pdfReader, pdfOutput);


                pdfStamper = new PdfStamper(pdfReader, new FileStream(TargetFile, FileMode.Create));
                AcroFields pdfFormFields = pdfStamper.AcroFields;
                foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
                {
                    if (de.Key.ToString() == TextName)
                    {
                        pdfFormFields.SetField(de.Key.ToString(), "");
                    }
                }
                //pdfReader.Close();

                //System.IO.File.WriteAllBytes(filePath, pdfOutput.ToArray());


            }
            catch (Exception ex)
            {
                return "{\"message\":\"Error\" " + ex.Message + "}";
            }
            finally
            {

                pdfStamper.Close();
                pdfReader.Close();
                System.IO.File.Copy(TargetFile, filePath, true);
            }

            return "{\"message\":\"Success\"}";
        }

        */

        public string SetFileNamePrefix()
        {
            
            return "";

        }

        public string SetTableName()
        {

            return "";

        }

        public IActionResult PreviewPdf(string SourceFile)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

            var uploadsFolder = Path.Combine(_environment.WebRootPath, "MapPdf/" + cmpid);
            var filePath = Path.Combine(uploadsFolder, SourceFile);
            // Check if the file exists
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            // Read the file and return as a FileStreamResult
            //var fileStream = System.IO.File.OpenRead(filePath);
            //return File(fileStream, "application/pdf");
            var fileStream = System.IO.File.OpenRead(filePath);
            return File(fileStream, "application/pdf");
        }

        [HttpPost]
        public IActionResult TransferPdf()
        {
                return View();
        }
    }
}
