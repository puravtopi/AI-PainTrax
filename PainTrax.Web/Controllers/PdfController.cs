using AutoMapper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Geom;
//using PdfSharp.Pdf.AcroForms;
//using PdfSharp.Pdf.IO;
//using PdfSharp.Pdf;
//using PdfSharp.Drawing;
//using DocumentFormat.OpenXml.Spreadsheet;
//using PdfSharp.Fonts;


using iTextSharp.text.pdf;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using MS.Services;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Diagnostics;
namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class PdfController : Controller
    {
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<PdfController> _logger;
        private readonly PatientService _patientservices = new PatientService();
        public PdfController(ILogger<PdfController> logger, IWebHostEnvironment environment)
        {
            _environment = environment;
            _logger = logger;
        }
        public IActionResult Test()
        {
            try
            {
                string filePath = _environment.WebRootPath + "/Demo_New.pdf";
                string imgPath = _environment.WebRootPath + "/test.png";
                // Create a memory stream to hold the modified PDF
                MemoryStream memoryStream = new MemoryStream();

                using (PdfReader reader = new PdfReader(filePath))
             //   using (PdfWriter writer = new PdfWriter(memoryStream))
                {
//                    PdfDocument pdfDoc = new PdfDocument(reader, writer);

                    // Access the first page
               //     PdfPage page = pdfDoc.GetFirstPage();

                    // Update textbox value
                    //PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                    //PdfFormField textbox = form.GetField("demo_fname");
                 //   textbox.SetValue("Test");

                    // Remove existing image and create a new textbox with the same bounds
                    //PdfArray annots = page.GetPdfObject().GetAsArray(PdfName.Annots);
                    //if (annots != null)
                    //{
                    //    foreach (PdfDictionary annotDict in annots)
                    //    {
                    //        if (PdfName.Widget.Equals(annotDict.GetAsName(PdfName.Subtype)))
                    //        {
                    //            PdfArray rect = annotDict.GetAsArray(PdfName.Rect);
                    //            Rectangle bbox = new Rectangle(rect.GetAsNumber(0).FloatValue(), rect.GetAsNumber(1).FloatValue(),
                    //                    rect.GetAsNumber(2).FloatValue(), rect.GetAsNumber(3).FloatValue());

                    //            // Replace image with textbox
                    //            annotDict.Remove(PdfName.AP);
                    //            annotDict.Put(PdfName.V, new PdfString("New Value")); // Set textbox value
                    //            annotDict.Put(PdfName.Subtype, PdfName.Widget);
                    //            annotDict.Put(PdfName.FT, PdfName.Tx);
                    //            annotDict.Put(PdfName.Rect, new PdfArray(bbox));
                    //        }
                    //    }
                    //}

                    // Close the document
          //          pdfDoc.Close();
                }

                // Reset the memory stream position to the beginning
                //memoryStream.Position = 0;

                // Return the modified PDF as a file
                byte[] fileBytes = memoryStream.ToArray();
                return File(fileBytes, "application/pdf", "output.pdf");
            }
            catch (Exception ex)
            {
                SaveLog(ex, "iText");
                return View();
            }

        }

        public IActionResult MapPdf()
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            
            string pdfFile = HttpContext.Session.GetString("SelectFile");
            var viewStructure = _patientservices.GetData("DESCRIBE View_Pdf;");
            ViewBag.ViewStructure = viewStructure;

            
            var filesByFolder = new Dictionary<string, List<string>>();
            var downloadFolder = System.IO.Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid.TrimEnd());
            var subFolders = Directory.GetDirectories(downloadFolder);
            List<string> fileList = new List<string>();

            foreach (var folder in subFolders)
            {
                fileList = new List<string>();
                var folderName = System.IO.Path.GetFileName(folder);
                var pdfFiles = Directory.GetFiles(folder, "*.pdf")
                                        .Select(System.IO.Path.GetFileName)
                                        .ToList();

                foreach (var item in pdfFiles)
                {
                    
                        fileList.Add(item);
                    
                }
                filesByFolder.Add(folderName, fileList);
            }

            ViewBag.FilesByFolder = filesByFolder;


            
            string webRoot = _environment.WebRootPath;
            string tempFile = System.IO.Path.Combine(webRoot, "temp.pdf");

            if(System.IO.File.Exists(tempFile))
            {
                PdfReader pdfReader = new PdfReader(tempFile);
                List<string> fieldNames = new List<string>();
                foreach (var fieldName in pdfReader.AcroFields.Fields.Keys)
                {
                    string fieldValue = pdfReader.AcroFields.GetField(fieldName);

                    // store as Name,Value
                    fieldNames.Add($"{fieldName},{fieldValue}");
                }
                
                ViewBag.Fields = fieldNames;
            }
            if (pdfFile != null)
            {
                ViewBag.PdfFile = pdfFile;
            }




            return View();
        }
        public IActionResult Index()
        {


            byte[] data = HttpContext.Session.Get("ExcelData");
            string pdfFile = HttpContext.Session.GetString("PdfFile");
            string excelFile = HttpContext.Session.GetString("ExcelFile");
            string fieldsJson = HttpContext.Session.GetString("PdfFields");
            string prefix = HttpContext.Session.GetString("FilePrefix");

            if (data != null)
            {

                DataSet ds = new DataSet();
                using (var ms = new MemoryStream(data))
                {
                    ds.ReadXml(ms);
                }

                // Take FIRST sheet (DataTable)
                DataTable dt = ds.Tables[0];
              
                // Send column names to ViewBag
                ViewBag.Columns = dt.Columns
                                    .Cast<DataColumn>()
                                    .Select(c => c.ColumnName)
                                    .ToList();
            }
            if(pdfFile != null)
            {
                ViewBag.PdfFile = pdfFile;  
            }
            if (excelFile != null)
            {
                ViewBag.excelFile = excelFile;
            }
            if (prefix != null)
            {
                ViewBag.prefix = prefix;
            }

            if (fieldsJson != null)
            {
                ViewBag.Fields  = System.Text.Json.JsonSerializer.Deserialize<List<string>>(fieldsJson);
            }

            return View();

            //string filePath = _environment.WebRootPath+ "/Demo_New.pdf";
            //string imgPath = _environment.WebRootPath+ "/test.png";
            //byte[] pdfBytes = System.IO.File.ReadAllBytes(filePath);
            //GlobalFontSettings.FontResolver = new MyFontResolver();

            //using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
            //{
            //    // Open existing PDF document from memory
            //    PdfDocument document = PdfReader.Open(pdfStream, PdfDocumentOpenMode.Modify);

            //    PdfAcroForm form = document.AcroForm;

            //    if (form != null)
            //    {

            //        // Iterate through each field in the form
            //        //  foreach (string fieldName in form.Fields.Names)
            //        foreach (var fieldName in form.Fields.Names)
            //        {
            //            // Get the field object by name
            //            //     PdfTextField textField = form.Fields. .FirstOrDefault(field => field.Name == fieldName);
            //            //  PdfAcroField field = fieldEntry.Value;


            //            PdfAcroField field = document.AcroForm.Fields[fieldName];
            //            // Check if the field is a textbox and its name matches the desired textbox

            //            if (field.Name == "demo_fname")
            //            {
            //                // Set the value of the textbox


            //                field.Value = new PdfString("New Value");

            //                // Get the first page of the document

            //            }

            //            if (field.Name == "demo_sex")
            //            {
            //                // Set the value of the textbox



            //                // Get the first page of the document
            //                PdfPage page = document.Pages[0]; // Assuming you want to add the image annotation to the first page

            //                // Define the position and size for the image annotation
            //                double imageX = field.Elements.GetRectangle("/Rect").X1 + 10; // Adjust as needed
            //                double imageY = field.Elements.GetRectangle("/Rect").Y1 + 10; // Adjust as needed
            //                double imageWidth = 50; // Adjust as needed
            //                double imageHeight = 50; // Adjust as needed

            //                // Add the image annotation to the page
            //                XGraphics gfx = XGraphics.FromPdfPage(page);
            //                XImage image = XImage.FromFile(imgPath); // Replace "image.jpg" with your image path
            //                gfx.DrawImage(image, imageX, imageY, imageWidth, imageHeight);
            //            }

            //        }
            //    }


            //    // Save the modified document to a new memory stream
            //    MemoryStream outputStream = new MemoryStream();
            //    document.Save(outputStream, false);
            //    byte[] updatedPdfBytes = outputStream.ToArray();
            //    return File(updatedPdfBytes, "application/pdf", "updated.pdf");

            //}

        }

        [HttpPost]
        public IActionResult UploadExcel(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
                return RedirectToAction("Index");

            DataSet ds = new DataSet();

            using (var stream = excelFile.OpenReadStream())
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart!;
                SharedStringTable sharedStringTable =
                    workbookPart.SharedStringTablePart?.SharedStringTable;

                foreach (Sheet sheet in workbookPart.Workbook.Sheets!)
                {
                    DataTable dt = new DataTable(sheet.Name!);

                    WorksheetPart worksheetPart =
                        (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    bool firstRow = true;

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (firstRow)
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                dt.Columns.Add(GetCellValue(cell, sharedStringTable));
                            }
                            firstRow = false;
                        }
                        else
                        {
                            DataRow dr = dt.NewRow();
                            int colIndex = 0;

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                dr[colIndex++] = GetCellValue(cell, sharedStringTable);
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                    ds.Tables.Add(dt);
                }
            }
            

            // Send column names to ViewBag
            TempData["Columns"] = ds.Tables[0].Columns
                                .Cast<DataColumn>()
                                .Select(c => c.ColumnName)
                                .ToList();

            // Store DataSet in Session
            HttpContext.Session.SetString("ExcelFile", excelFile.FileName);
            HttpContext.Session.Set("ExcelData", SerializeDataSet(ds));
            //byte[] data = HttpContext.Session.Get("ExcelData");

            //DataSet ds = new DataSet();
            //using var ms = new MemoryStream(data);
            //ds.ReadXml(ms);
            return RedirectToAction("Index");
        }


        private static string GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
        {
            if (cell.CellValue == null) return "";

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedStringTable!.ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }

        private byte[] SerializeDataSet(DataSet ds)
        {
            using var ms = new MemoryStream();
            ds.WriteXml(ms, XmlWriteMode.WriteSchema);
            return ms.ToArray();
        }
        [HttpPost]
        public IActionResult UploadPdf(IFormFile pdfFile)
        {
            if (pdfFile == null || pdfFile.Length == 0)
                return RedirectToAction("Index");

            string webRoot = _environment.WebRootPath;
            string tempPath = System.IO.Path.Combine(webRoot, "_temp.pdf");

            // Save uploaded pdf as _temp.pdf
            using (var fs = new FileStream(tempPath, FileMode.Create))
            {
                pdfFile.CopyTo(fs);
            }
            PdfReader pdfReader = new PdfReader(tempPath);
            List<string> fieldNames = new List<string>();
            fieldNames = pdfReader.AcroFields.Fields.Keys.ToList();
            HttpContext.Session.SetString("PdfFile", pdfFile.FileName);
            HttpContext.Session.SetString("PdfFields",System.Text.Json.JsonSerializer.Serialize(fieldNames));
            return RedirectToAction("Index");
        }

        [HttpPost]        
        public IActionResult MapField(string excelColumn, string pdfFieldName)
        {
            string webRoot = _environment.WebRootPath;
            //string tempPath = System.IO.Path.Combine(webRoot, "_temp.pdf");
            string inputPdf = System.IO.Path.Combine(webRoot, "_temp.pdf");
            string outputPdf = System.IO.Path.Combine(webRoot, "_temp2.pdf");

            using (PdfReader reader = new PdfReader(inputPdf))
            using (FileStream fs = new FileStream(outputPdf, FileMode.Create, FileAccess.Write))
            using (PdfStamper stamper = new PdfStamper(reader, fs))
            {
                AcroFields fields = stamper.AcroFields;

                // Rename PDF textbox → #ExcelColumn
                string newName = "#" + excelColumn;

                if (fields.GetField(pdfFieldName) != null)
                {
                    fields.RenameField(pdfFieldName, newName);
                }
            }
            System.IO.File.Delete(inputPdf);
            System.IO.File.Move(outputPdf, inputPdf);
            PdfReader pdfReader = new PdfReader(inputPdf);
            List<string> fieldNames = new List<string>();
            fieldNames = pdfReader.AcroFields.Fields.Keys.ToList();
            HttpContext.Session.SetString("PdfFields", System.Text.Json.JsonSerializer.Serialize(fieldNames));

            return Json(new { message = "PDF field mapped successfully" });
        }

        [HttpPost]
        public IActionResult ClearExcel()
        {
            HttpContext.Session.Remove("ExcelData");
            HttpContext.Session.Remove("ExcelFile");
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult ClearPdf()
        {
            HttpContext.Session.Remove("PdfFile");
            HttpContext.Session.Remove("PdfFields");
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult GenrateZip()
        {
            string webRoot = _environment.WebRootPath;
            byte[] data = HttpContext.Session.Get("ExcelData");
            string pdfFile = HttpContext.Session.GetString("PdfFile");
            string excelFile = HttpContext.Session.GetString("ExcelFile");
            string fieldsJson = HttpContext.Session.GetString("PdfFields");
            string inputPdf = System.IO.Path.Combine(webRoot, "_temp.pdf");
            string prefix = HttpContext.Session.GetString("FilePrefix");

            DataTable dt = new DataTable();
            if (data != null)
            {

                DataSet ds = new DataSet();
                using (var ms = new MemoryStream(data))
                {
                    ds.ReadXml(ms);
                }

                // Take FIRST sheet (DataTable)
                dt = ds.Tables[0];
            }
            string tempFolder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempFolder);

            foreach (DataRow row in dt.Rows)
            {
                //string lname = row["LName"]?.ToString();
                //string fname = row["FName"]?.ToString();

                //string pdfName = $"{lname}_{fname}_{pdfFile}.pdf";
                List<string> nameParts = new List<string>();

                if (!string.IsNullOrWhiteSpace(prefix))
                {
                    // extract column names between [ ]
                    var matches = System.Text.RegularExpressions.Regex.Matches(prefix, @"\[(.*?)\]");

                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        string columnName = match.Groups[1].Value;

                        // check column exists and value not null
                        if (dt.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                        {
                            string value = row[columnName].ToString();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                nameParts.Add(value);
                            }
                        }
                    }
                }

                // join parts with underscore
                string dynamicName = string.Join(",", nameParts);

                // final pdf name
                string pdfName = string.IsNullOrEmpty(dynamicName)
                    ? $"{pdfFile}.pdf"
                    : $"{dynamicName}_{pdfFile}.pdf";
                string outputPdf = System.IO.Path.Combine(tempFolder, pdfName);

                GeneratePdf(inputPdf, outputPdf, row);
            }

            // Create ZIP
            string zipPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"PDFs_{DateTime.Now.Ticks}.zip");
            System.IO.Compression.ZipFile.CreateFromDirectory(tempFolder, zipPath);

            byte[] zipBytes = System.IO.File.ReadAllBytes(zipPath);

            // Cleanup
            Directory.Delete(tempFolder, true);
            System.IO.File.Delete(zipPath);

            return File(zipBytes, "application/zip", "GeneratedPDFs.zip");

            //return RedirectToAction("Index");
        }

        private void GeneratePdf(string templatePath, string outputPath, DataRow row)
        {
            using (PdfReader reader = new PdfReader(templatePath))
            using (FileStream fs = new FileStream(outputPath, FileMode.Create))
            using (PdfStamper stamper = new PdfStamper(reader, fs))
            {
                AcroFields fields = stamper.AcroFields;

                foreach (DataColumn col in row.Table.Columns)
                {
                    string fieldName = "#" + col.ColumnName;
                    string value = row[col]?.ToString();

                    if (fields.Fields.ContainsKey(fieldName))
                    {
                        fields.SetField(fieldName, value);
                    }
                }

                stamper.FormFlattening = true; // optional (make PDF non-editable)
            }
        }

        [HttpPost]
        public IActionResult DownloadMapPdf()
        {
            // get web root path
            string webRoot = _environment.WebRootPath;

            // path of generated pdf
            string tempPdfPath = System.IO.Path.Combine(webRoot, "_temp.pdf");

            if (!System.IO.File.Exists(tempPdfPath))
            {
                return NotFound("PDF file not found.");
            }

            // get original pdf filename from session
            string pdfFile = HttpContext.Session.GetString("PdfFile");

            // fallback name if session is empty
            if (string.IsNullOrEmpty(pdfFile))
            {
                pdfFile = "MappedPdf.pdf";
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(tempPdfPath);

            return File(
                fileBytes,
                "application/pdf",
                pdfFile
            );
        }

        [HttpPost]
        public IActionResult SetFilePrefix(string prefix)
        {
            HttpContext.Session.SetString("FilePrefix", prefix);
            return Json(new { success = true });
        }

        [HttpPost]
        public IActionResult SetPdf(string folderName, string pdfName)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();

            if (string.IsNullOrWhiteSpace(pdfName))
            {
                TempData["Error"] = "Please select a PDF file.";
                return RedirectToAction("MapPdf");
            }


            
            string downloadFolder = System.IO.Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid.TrimEnd());
            string sourceFile = System.IO.Path.Combine(downloadFolder, folderName, pdfName);

            if (!System.IO.File.Exists(sourceFile))
            {
                TempData["Error"] = "Selected PDF file not found.";
                return RedirectToAction("MapPdf");
            }

            string webRoot = _environment.WebRootPath;
            string destinationFile = System.IO.Path.Combine(webRoot, "temp.pdf");
            HttpContext.Session.SetString("SelectFile", folderName + "/"+  pdfName);

            System.IO.File.Copy(sourceFile, destinationFile, true);

            return RedirectToAction("MapPdf");
        }

        [HttpPost]        
        public IActionResult ClearMapPdf()
        {
            string webRoot = _environment.WebRootPath;
            string destinationFile = System.IO.Path.Combine(webRoot, "temp.pdf");
            System.IO.File.Delete(destinationFile);
            HttpContext.Session.Remove("SelectFile");
            return RedirectToAction("MapPdf");
        }

            [HttpPost]
        public IActionResult MapTran(string col, string text, string type, string settext)
        {
            string webRoot = _environment.WebRootPath;
            //string tempPath = System.IO.Path.Combine(webRoot, "_temp.pdf");
            string inputPdf = System.IO.Path.Combine(webRoot, "temp.pdf");
            string outputPdf = System.IO.Path.Combine(webRoot, "temp2.pdf");


            string fieldname = text.Split(",")[0];
            using (PdfReader reader = new PdfReader(inputPdf))
            using (FileStream fs = new FileStream(outputPdf, FileMode.Create, FileAccess.Write))
            using (PdfStamper stamper = new PdfStamper(reader, fs))
            {
                AcroFields fields = stamper.AcroFields;

                if (type == "SetText" )
                {

                    if (fields.GetField(fieldname) != null)
                    {
                        // Set value AFTER rename
                        fields.SetField(fieldname, settext);
                    }
                }

                if (type == "Clear" || type == "MapClear")
                {

                    if (fields.GetField(fieldname) != null)
                    {
                        // Set value AFTER rename
                        fields.SetField(fieldname, "");
                    }
                }
                // Rename PDF textbox → #ExcelColumn
                string newName = "#" + col;
                if (type == "Map" || type == "MapClear")
                {
                    if (fields.GetField(fieldname) != null)
                    {
                        fields.RenameField(fieldname, newName);
                    }
                }
            }
            
            System.IO.File.Delete(inputPdf);
            System.IO.File.Move(outputPdf, inputPdf);
            PdfReader pdfReader = new PdfReader(inputPdf);
            List<string> fieldNames = new List<string>();
            fieldNames = pdfReader.AcroFields.Fields.Keys.ToList();
            HttpContext.Session.SetString("PdfFields", System.Text.Json.JsonSerializer.Serialize(fieldNames));

            return Json(new { message = "PDF field mapped successfully" });
        }

        [HttpPost]
        public IActionResult SaveMapPdf()
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();

            string webRoot = _environment.WebRootPath;
            string tempFile = System.IO.Path.Combine(webRoot, "temp.pdf");
                 string downloadFolder = System.IO.Path.Combine(_environment.WebRootPath, "Downloads/" + cmpclientid.TrimEnd());
            
            if(HttpContext.Session.GetString("SelectFile")!=null || HttpContext.Session.GetString("SelectFile")!="")
            {
                string sourceFile = System.IO.Path.Combine(downloadFolder, HttpContext.Session.GetString("SelectFile"));
                System.IO.File.Copy(tempFile,sourceFile,true);
            }
            return RedirectToAction("MapPdf");
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
//    public class MyFontResolver : IFontResolver
//    {
//        public byte[] GetFont(string faceName)
//        {
//            // Implement your font retrieval logic here
//            // This method is not used if you want to ignore a font, so you can leave it unchanged
//            return null;
//        }

//        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
//        {
//            // Implement your logic to resolve typefaces here
//            // You can return default font information for specific font families you want to ignore

//            // Example: Ignore "Courier New" font and use the default font instead
//            if (familyName.Equals("Courier New", StringComparison.OrdinalIgnoreCase))
//            {
//                // Return the default font information (e.g., Arial) for "Courier New"
//                return new FontResolverInfo("Arial");
//            }

//            // Return null if you want to use PDFSharp's default font resolution logic for other font families
//            return null;
//        }
//    }

//}
