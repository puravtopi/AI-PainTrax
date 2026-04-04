using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;


namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class ProcedureController : Controller
    {
        private readonly IMapper _mapper;
        private readonly ILogger<ProcedureController> _logger;
        private readonly ProcedureService _services = new ProcedureService();
        private readonly Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;

        public ProcedureController(IMapper mapper, ILogger<ProcedureController> logger, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _mapper = mapper;
            _logger = logger;
            Environment = environment;
        }

        public IActionResult Index(string searchtxt = "")
        {
            return View();
        }

        public IActionResult Create()
        {
            tbl_procedures obj = new tbl_procedures();
            try
            {
                obj.inhouseprocbit = false;
                obj.sides = false;
                obj.haslevel = false;
                obj.cf = false;
                obj.pn = false;
                obj.inout = false;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return View(obj);
        }


        [HttpPost]
        [AllowAnonymous]
        public IActionResult Create(tbl_procedures model, IFormFile upload_template)
        {
            try
            {
                model.position = model.position == null ? "" : model.position;
                model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                if (upload_template != null)
                {
                    string folderPath = Path.Combine(Environment.WebRootPath, "Uploads/InjectionReports", model.cmp_id.ToString());
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    // string fileName = Path.GetFileName(upload_template.FileName);                    
                    //string filePath = Path.Combine(folderPath, fileName);
                    string originalFileName = Path.GetFileName(upload_template.FileName);
                    string extension = Path.GetExtension(originalFileName);
                    string fileNameasmcode = model.mcode;
                    string fileName = fileNameasmcode + extension;
                    string filePath = Path.Combine(folderPath, fileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        upload_template.CopyTo(fileStream);
                    }
                    model.upload_template = fileName;
                }
                _services.Insert(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Edit(int id)
        {
            var data = new tbl_procedures();
            try
            {
                tbl_procedures obj = new tbl_procedures();
                obj.id = id;
                data = _services.GetOne(obj);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_procedures model, IFormFile upload_template)
        {
            try
            {
                model.position = model.position == null ? "" : model.position;
                if (upload_template != null && upload_template.Length > 0)
                {
                    string folderPath = Path.Combine(Environment.WebRootPath, "Uploads/InjectionReports", model.cmp_id.ToString());
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string originalFileName = Path.GetFileName(upload_template.FileName);
                    string extension = Path.GetExtension(originalFileName);
                    string fileNameasmcode = model.mcode;
                    string fileName = fileNameasmcode + extension;
                    string filePath = Path.Combine(folderPath, fileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        upload_template.CopyTo(fileStream);
                    }
                    model.upload_template = fileName;
                }
                _services.Update(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Delete(int id)
        {
            try
            {
                tbl_procedures obj = new tbl_procedures();
                obj.id = id;
                _services.Delete(obj);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Delete");
            }
            return RedirectToAction("Index");
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
                var sortColumn = Request.Form["order[0][column]"].FirstOrDefault();
                // Sort Column Direction ( asc ,desc)
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                // Search Value from (Search box)
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                //Paging Size (10,20,50,100)
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                string cnd = " and cmp_id=" + cmpid + " and (mcode like '%" + searchValue + "%' or bodypart like '%" + searchValue + "%' or heading like '%" + searchValue + "%' or ccdesc like '%" + searchValue + "%' Or ";
                cnd += " pedesc like '%" + searchValue + "%' or adesc like '%" + searchValue + "%' or pdesc like '%" + searchValue + "%' or cf like '%" + searchValue + "%' )";

                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_procedures).GetProperties()[Convert.ToInt32(sortColumn)];
                    if (sortColumnDirection.ToUpper() == "ASC")
                    {
                        Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                    }
                    else
                    {
                        Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                    }
                }
                //Search
                //total number of rows count 
                recordsTotal = Data.Count();
                //Paging 
                var data = Data.Skip(skip).Take(pageSize).ToList();
                //Returning Json Data
                return Json(new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data });
            }
            catch (Exception ex)
            {
                SaveLog(ex, "List");
            }
            return Json("");
        }

        public IActionResult ImportData(IFormFile postedFile)
        {
            try
            {
                if (postedFile != null && postedFile.Length > 0)
                {
                    //DataTable dt = this.Read2007Xlsx(postedFile);
                    //DataTable dt = this.Read2007Xlsx(patient);
                    DataTable dt = new DataTable();
                    using (var stream = new MemoryStream())
                    {
                        postedFile.CopyToAsync(stream);
                        stream.Position = 0;

                        // Convert uploaded Excel to DataTable
                        dt = ReadExcelToDataTable(stream);
                    }
                    int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        tbl_procedures obj = new tbl_procedures();

                        obj.cmp_id = cmpid;
                        obj.position = (dt.Columns.Contains("Position") && dt.Rows[i]["Position"] != DBNull.Value ? dt.Rows[i]["Position"].ToString() : "");
                        obj.display_order = (dt.Columns.Contains("DisplayOrder") && dt.Rows[i]["DisplayOrder"] != DBNull.Value ? Convert.ToInt16(dt.Rows[i]["DisplayOrder"].ToString()) : 0);
                        obj.hasmuscle = (dt.Columns.Contains("Muscle") && dt.Rows[i]["Muscle"] != DBNull.Value ? dt.Rows[i]["Muscle"].ToString() : "");
                        obj.hassubprocedure = (dt.Columns.Contains("SubProcedure") && dt.Rows[i]["SubProcedure"] != DBNull.Value ? dt.Rows[i]["SubProcedure"].ToString() : "");
                        obj.pedesc = (dt.Columns.Contains("R_PEDesc") && dt.Rows[i]["R_PEDesc"] != DBNull.Value ? dt.Rows[i]["R_PEDesc"].ToString().Replace("'s", "") : "");
                        obj.pdesc = (dt.Columns.Contains("R_PDesc") && dt.Rows[i]["R_PDesc"] != DBNull.Value ? dt.Rows[i]["R_PDesc"].ToString().Replace("'s", "") : "");
                        obj.s_ccdesc = (dt.Columns.Contains("S_CCDesc") && dt.Rows[i]["S_CCDesc"] != DBNull.Value ? dt.Rows[i]["S_CCDesc"].ToString().Replace("'s", "") : ""); 
                        obj.s_adesc = (dt.Columns.Contains("S_ADesc") && dt.Rows[i]["S_ADesc"] != DBNull.Value ? dt.Rows[i]["S_ADesc"].ToString().Replace("'s", "") : "");
                        obj.e_heading = (dt.Columns.Contains("E_Heading") && dt.Rows[i]["E_Heading"] != DBNull.Value ? dt.Rows[i]["E_Heading"].ToString().Replace("'s", "") : ""); 
                        obj.e_pedesc = (dt.Columns.Contains("E_PEDesc") && dt.Rows[i]["E_PEDesc"] != DBNull.Value ? dt.Rows[i]["E_PEDesc"].ToString().Replace("'s", "") : "");
                        obj.e_pdesc = (dt.Columns.Contains("E_PDesc") && dt.Rows[i]["E_PDesc"] != DBNull.Value ? dt.Rows[i]["E_PDesc"].ToString().Replace("'s", "") : "");
                        obj.sidesdefault = (dt.Columns.Contains("SidesDefault") && dt.Rows[i]["SidesDefault"] != DBNull.Value ? dt.Rows[i]["SidesDefault"].ToString() : ""); 
                        obj.inhouseprocbit = (dt.Columns.Contains("InHouseProcbit") && dt.Rows[i]["InHouseProcbit"] != DBNull.Value ? ConvertToBoolean(dt.Rows[i]["InHouseProcbit"]) : false);
                        obj.sides = (dt.Columns.Contains("sides") && dt.Rows[i]["sides"] != DBNull.Value ? ConvertToBoolean(dt.Rows[i]["sides"]) : false);
                        obj.haslevel = (dt.Columns.Contains("haslevel") && dt.Rows[i]["haslevel"] != DBNull.Value ? ConvertToBoolean(dt.Rows[i]["haslevel"]) : false);
                        obj.cf = (dt.Columns.Contains("cf") && dt.Rows[i]["cf"] != DBNull.Value ? ConvertToBoolean(dt.Rows[i]["cf"]) : false); 
                        obj.pn = (dt.Columns.Contains("pn") && dt.Rows[i]["pn"] != DBNull.Value ? ConvertToBoolean(dt.Rows[i]["pn"]) : false); 
                        obj.inout = (dt.Columns.Contains("inout") && dt.Rows[i]["inout"] != DBNull.Value ? ConvertToBoolean(dt.Rows[i]["inout"]) : false); 
                        obj.bodypart = (dt.Columns.Contains("BodyPart") && dt.Rows[i]["BodyPart"] != DBNull.Value ? dt.Rows[i]["BodyPart"].ToString() : "");
                        obj.heading = (dt.Columns.Contains("Heading") && dt.Rows[i]["Heading"] != DBNull.Value ? dt.Rows[i]["Heading"].ToString() : "");
                        obj.hasmedication = (dt.Columns.Contains("HasMedication") && dt.Rows[i]["HasMedication"] != DBNull.Value ? dt.Rows[i]["HasMedication"].ToString() : ""); 
                        obj.ccdesc = (dt.Columns.Contains("CCDesc") && dt.Rows[i]["CCDesc"] != DBNull.Value ? dt.Rows[i]["CCDesc"].ToString().Replace("'s", "") : ""); 
                        obj.adesc = (dt.Columns.Contains("ADesc") && dt.Rows[i]["ADesc"] != DBNull.Value ? dt.Rows[i]["ADesc"].ToString().Replace("'s", "") : "");
                        obj.s_heading = (dt.Columns.Contains("S_Heading") && dt.Rows[i]["S_Heading"] != DBNull.Value ? dt.Rows[i]["S_Heading"].ToString().Replace("'s", "") : ""); 
                        obj.s_pedesc = (dt.Columns.Contains("S_PEDesc") && dt.Rows[i]["S_PEDesc"] != DBNull.Value ? dt.Rows[i]["S_PEDesc"].ToString().Replace("'s", "") : "");
                        obj.e_ccdesc = (dt.Columns.Contains("E_CCDesc") && dt.Rows[i]["E_CCDesc"] != DBNull.Value ? dt.Rows[i]["E_CCDesc"].ToString().Replace("'s", "") : "");
                        obj.e_adesc = (dt.Columns.Contains("E_ADesc") && dt.Rows[i]["E_ADesc"] != DBNull.Value ? dt.Rows[i]["E_ADesc"].ToString().Replace("'s", "") : ""); 
                        obj.levelsdefault = (dt.Columns.Contains("LevelsDefault") && dt.Rows[i]["LevelsDefault"] != DBNull.Value ? dt.Rows[i]["LevelsDefault"].ToString() : "");
                        obj.mcode = (dt.Columns.Contains("MCode") && dt.Rows[i]["MCode"] != DBNull.Value ? dt.Rows[i]["MCode"].ToString() : "");
                        obj.mcode_desc = (dt.Columns.Contains("MCodeDesc") && dt.Rows[i]["MCodeDesc"] != DBNull.Value ? dt.Rows[i]["MCodeDesc"].ToString() : "");
                        obj.s_pdesc = (dt.Columns.Contains("S_PDesc") && dt.Rows[i]["S_PDesc"] != DBNull.Value ? dt.Rows[i]["S_PDesc"].ToString() : "");
                       

                        _services.Insert(obj);


                    }
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "ImportData");
            }
            return RedirectToAction("Index");
        }
        private bool ConvertToBoolean(object value)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString()))
                return false;

            string val = value.ToString().ToLower();
            return val == "1" || val == "true" || val == "yes";
        }

        //public DataTable Read2007Xlsx(IFormFile postedFile)
        //{
        //    DataTable dt = new DataTable();
        //    try
        //    {
        //        if (postedFile != null && postedFile.Length > 0)
        //        {
        //            // Read the uploaded Excel file using Open XML
        //            using (Stream stream = postedFile.OpenReadStream())
        //            {

        //                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(stream, false))
        //                {
        //                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
        //                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
        //                    string relationshipId = sheets.First().Id.Value;
        //                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
        //                    Worksheet workSheet = worksheetPart.Worksheet;
        //                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
        //                    IEnumerable<Row> rows = sheetData.Descendants<Row>();
        //                    foreach (Cell cell in rows.ElementAt(0))
        //                    {
        //                        dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
        //                    }
        //                    foreach (Row row in rows) //this will also include your header row...
        //                    {
        //                        DataRow tempRow = dt.NewRow();
        //                        int columnIndex = 0;
        //                        foreach (Cell cell in row.Descendants<Cell>())
        //                        {
        //                            // Gets the column index of the cell with data
        //                            int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
        //                            cellColumnIndex--; //zero based index
        //                            if (columnIndex < cellColumnIndex)
        //                            {
        //                                do
        //                                {
        //                                    tempRow[columnIndex] = ""; //Insert blank data here;
        //                                    columnIndex++;
        //                                }
        //                                while (columnIndex < cellColumnIndex);
        //                            }//end if block
        //                            tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);
        //                            columnIndex++;
        //                        }//end inner foreach loop
        //                        dt.Rows.Add(tempRow);
        //                    }//end outer foreach loop
        //                }//end using block
        //                dt.Rows.RemoveAt(0); //...so i'm taking it out here.
        //            }
        //        }
        //    }//end try
        //    catch (Exception ex)
        //    {
        //        SaveLog(ex, "Read2007Xlsx");
        //    }

        //    return dt;
        //}//end Read2007Xlsx method        
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

        [HttpGet]
        public ActionResult DownloadDocument()
        {
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "Procedure.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "Procedure.xlsx");
        }

        [HttpGet]
        public IActionResult ExportToExcel()
        {
            try
            {
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                string cnd = " and cmp_id=" + cmpid;

                var data = _services.GetAll(cnd); // Retrieve all attorneys from the database

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {
                    new DataColumn("Position", typeof(string)),
                    new DataColumn("DisplayOrder", typeof(int)),
                    new DataColumn("Muscle", typeof(string)),
                    new DataColumn("Sub Procedure", typeof(string)),
                    new DataColumn("R_PEDesc", typeof(string)),
                    new DataColumn("R_PDesc", typeof(string)),
                    new DataColumn("S_CCDesc", typeof(string)),
                    new DataColumn("S_ADesc", typeof(string)),
                    new DataColumn("E-Heading", typeof(string)),
                    new DataColumn("E_PEDesc", typeof(string)),
                    new DataColumn("E_PDesc", typeof(string)),
                    new DataColumn("SidesDefault", typeof(string)),
                    new DataColumn("InHouseProcbit", typeof(bool)),
                    new DataColumn("sides", typeof(bool)),
                    new DataColumn("haslevel", typeof(bool)),
                    new DataColumn("cf", typeof(bool)),
                    new DataColumn("pn", typeof(bool)),
                    new DataColumn("InOut", typeof(bool)),
                    new DataColumn("BodyPart", typeof(string)),
                    new DataColumn("Heading", typeof(string)),
                    new DataColumn("HasMedication", typeof(string)),
                    new DataColumn("CCDesc", typeof(string)),
                    new DataColumn("ADesc", typeof(string)),
                    new DataColumn("S_PEDesc", typeof(string)),
                    new DataColumn("S_Heading", typeof(string)),
                    new DataColumn("E_CCDesc", typeof(string)),
                    new DataColumn("E_ADesc", typeof(string)),
                    new DataColumn("LevelsDefault", typeof(string)),
                    new DataColumn("MCode", typeof(string)),
                    new DataColumn("MCode Desc", typeof(string))
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var procedure in data)
                {
                    dt.Rows.Add(procedure.position, procedure.display_order, procedure.hasmuscle, procedure.hassubprocedure, procedure.pedesc, procedure.pdesc, procedure.s_ccdesc, procedure.s_adesc,
                        procedure.e_heading, procedure.e_pedesc, procedure.e_pdesc, procedure.sidesdefault, procedure.inhouseprocbit, procedure.sides, procedure.haslevel, procedure.cf, procedure.pn,
                        procedure.inout, procedure.bodypart, procedure.heading, procedure.hasmedication, procedure.ccdesc, procedure.adesc, procedure.s_pedesc, procedure.s_heading, procedure.e_ccdesc,
                        procedure.e_adesc, procedure.levelsdefault, procedure.mcode, procedure.mcode_desc);
                }

                // Create a new Excel file
                var memoryStream = new MemoryStream();
                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Procedure" };
                    sheets.Append(sheet);

                    var headerRow = new Row();
                    foreach (DataColumn column in dt.Columns)
                    {
                        var cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(column.ColumnName) };
                        headerRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow row in dt.Rows)
                    {
                        var newRow = new Row();
                        foreach (var value in row.ItemArray)
                        {
                            var cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(value.ToString()) };
                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }
                }

                memoryStream.Seek(0, SeekOrigin.Begin);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Procedure.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }

        //public IActionResult Clone(int id)
        //{
        //    try
        //    {
        //        // Fetch the procedure to be cloned
        //        var eProcedure = _services.GetOne(new tbl_procedures { id = id });
        //        if (eProcedure == null)
        //        {
        //            return NotFound("Procedure not found.");
        //        }

        //        // Create a new procedure object with the same data
        //        var clonedProcedure = new tbl_procedures
        //        {
        //            cmp_id = eProcedure.cmp_id,
        //            position = eProcedure.position,
        //            display_order = eProcedure.display_order,
        //            hasmuscle = eProcedure.hasmuscle,
        //            hassubprocedure = eProcedure.hassubprocedure,
        //            pedesc = eProcedure.pedesc,
        //            pdesc = eProcedure.pdesc,
        //            s_ccdesc = eProcedure.s_ccdesc,
        //            s_adesc = eProcedure.s_adesc,
        //            e_heading = eProcedure.e_heading,
        //            e_pedesc = eProcedure.e_pedesc,
        //            e_pdesc = eProcedure.e_pdesc,
        //            sidesdefault = eProcedure.sidesdefault,
        //            inhouseprocbit = eProcedure.inhouseprocbit,
        //            sides = eProcedure.sides,
        //            haslevel = eProcedure.haslevel,
        //            cf = eProcedure.cf,
        //            pn = eProcedure.pn,
        //            inout = eProcedure.inout,
        //            bodypart = eProcedure.bodypart,
        //            heading = eProcedure.heading,
        //            hasmedication = eProcedure.hasmedication,
        //            ccdesc = eProcedure.ccdesc,
        //            adesc = eProcedure.adesc,
        //            s_heading = eProcedure.s_heading,
        //            s_pedesc = eProcedure.s_pedesc,
        //            e_ccdesc = eProcedure.e_ccdesc,
        //            e_adesc = eProcedure.e_adesc,
        //            levelsdefault = eProcedure.levelsdefault,
        //            mcode = eProcedure.mcode,
        //            mcode_desc = eProcedure.mcode_desc,
        //            upload_template = eProcedure.upload_template // Copy the template file name
        //        };

        //        _services.Insert(clonedProcedure);

        //        return RedirectToAction("Index", new { id = clonedProcedure.id });
        //    }
        //    catch (Exception ex)
        //    {
        //        SaveLog(ex, "Clone");
        //        return RedirectToAction("Index");
        //    }
        //}
        [HttpPost]
        public IActionResult Clone(int id)
        {
            try
            {
                var originalProcedure = _services.GetOne(new tbl_procedures { id = id });
                if (originalProcedure != null)
                {
                    var clonedProcedure = new tbl_procedures
                    {
                        bodypart = originalProcedure.bodypart,
                        mcode = originalProcedure.mcode,
                        heading = originalProcedure.heading,
                        cmp_id = originalProcedure.cmp_id,
                        position = originalProcedure.position,
                        display_order = originalProcedure.display_order,
                        hasmuscle = originalProcedure.hasmuscle,
                        hassubprocedure = originalProcedure.hassubprocedure,
                        pedesc = originalProcedure.pedesc,
                        pdesc = originalProcedure.pdesc,
                        s_ccdesc = originalProcedure.s_ccdesc,
                        s_adesc = originalProcedure.s_adesc,
                        e_heading = originalProcedure.e_heading,
                        e_pedesc = originalProcedure.e_pedesc,
                        e_pdesc = originalProcedure.e_pdesc,
                        sidesdefault = originalProcedure.sidesdefault,
                        inhouseprocbit = originalProcedure.inhouseprocbit,
                        sides = originalProcedure.sides,
                        haslevel = originalProcedure.haslevel,
                        cf = originalProcedure.cf,
                        pn = originalProcedure.pn,
                        inout = originalProcedure.inout,
                        hasmedication = originalProcedure.hasmedication,
                        ccdesc = originalProcedure.ccdesc,
                        adesc = originalProcedure.adesc,
                        s_heading = originalProcedure.s_heading,
                        s_pedesc = originalProcedure.s_pedesc,
                        e_ccdesc = originalProcedure.e_ccdesc,
                        e_adesc = originalProcedure.e_adesc,
                        levelsdefault = originalProcedure.levelsdefault,
                        mcode_desc = originalProcedure.mcode_desc,
                        upload_template = originalProcedure.upload_template
                    };

                    _services.Insert(clonedProcedure);
                    return Json(new { success = true, message = "Procedure clone successfully" });
                }
                else
                {
                    return Json(new { success = false, message = "Original procedure not found." });
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Json(new { success = false, message = "An error occurred while cloning the procedure." });
            }
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

                    /*                   foreach (Cell cell in row.Elements<Cell>())
                                       {
                                           Console.WriteLine(cell.Count());
                                           string cellValue = GetCellValue(spreadsheetDocument, cell);

                                           if (isFirstRow)
                                           {
                                               // Use the first row to add columns to the DataTable
                                               dataTable.Columns.Add(cellValue);
                                           }
                                           else
                                           {
                                               // Clean HTML tags and add data to the DataTable
                                               string cleanCellValue = cellValue; //System.Text.RegularExpressions.Regex.Replace(cellValue, "<.*?>", string.Empty);
                                               dataRow[columnIndex] = cleanCellValue;
                                           }
                                           columnIndex++;
                                       }

                                       if (!isFirstRow)
                                       {
                                           dataTable.Rows.Add(dataRow);
                                       }
                                       isFirstRow = false;

                    */

                    int currentColumnIndex = 0;
                    int previousColumnIndex = -1;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        // Get the column index from the cell reference (e.g., A1, B1)
                        string cellReference = cell.CellReference?.Value;
                        int cellColumnIndex = GetColumnIndexFromCellReference(cellReference);

                        // Check for missing cells by comparing current cell index with previous one
                        if (cellColumnIndex - previousColumnIndex > 1)
                        {
                            // Fill missing columns with null
                            int missingCells = cellColumnIndex - previousColumnIndex - 1;
                            for (int i = 0; i < missingCells; i++)
                            {
                                if (isFirstRow)
                                {
                                    // Add placeholder column headers for missing columns
                                    dataTable.Columns.Add($"Column {currentColumnIndex + 1}");
                                }
                                else
                                {
                                    // Add null for missing cells in the data row
                                    dataRow[currentColumnIndex] = DBNull.Value;
                                }
                                currentColumnIndex++;
                            }
                        }

                        // Process current cell value
                        string cellValue = GetCellValue(spreadsheetDocument, cell);

                        if (isFirstRow)
                        {
                            // Add the column header for non-empty cells
                            dataTable.Columns.Add(cellValue ?? $"Column {currentColumnIndex + 1}");
                        }
                        else
                        {
                            // Clean HTML tags and add data to the DataTable
                            string cleanCellValue = System.Text.RegularExpressions.Regex.Replace(cellValue, "<.*?>", string.Empty);
                            dataRow[currentColumnIndex] = string.IsNullOrEmpty(cleanCellValue) ? DBNull.Value : cleanCellValue;
                        }

                        previousColumnIndex = cellColumnIndex;
                        currentColumnIndex++;
                    }

                    // Fill any remaining columns with nulls if the row is shorter than the header
                    while (currentColumnIndex < dataTable.Columns.Count)
                    {
                        dataRow[currentColumnIndex] = DBNull.Value;
                        currentColumnIndex++;
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

        private int GetColumnIndexFromCellReference(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return -1;
            }

            // Extract the column part (e.g., "A" from "A1")
            string columnPart = new string(cellReference.Where(c => char.IsLetter(c)).ToArray());

            int columnIndex = 0;
            foreach (char c in columnPart)
            {
                columnIndex = (columnIndex * 26) + (c - 'A' + 1);
            }

            return columnIndex - 1; // Zero-based index
        }
        private string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell == null)
                return "";

            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                value = stringTable.ChildElements[int.Parse(value)].InnerText;
            }


            return value;
        }
        #endregion
    }
}
