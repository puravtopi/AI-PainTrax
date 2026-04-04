using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class DiagcodeController : Controller
    {
        private readonly IMapper _mapper;
        private readonly DiagcodesService _services = new DiagcodesService();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;
        private readonly ILogger<DiagcodeController> _logger;

        public DiagcodeController(IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
            IConfiguration configuration, ILogger<DiagcodeController> logger)
        {
            _logger = logger;
            _mapper = mapper;
            Environment = environment;
            Configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Create()
        {
            tbl_diagcodes obj = new tbl_diagcodes();
            // obj.PreSelect = false;
            return View(obj);
        }


        [HttpPost]
        [AllowAnonymous]
        public IActionResult Create(tbl_diagcodes model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    model.CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.CreatedDate = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    //model.PreSelect =null;
                    _services.Insert(model);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Edit(int id)
        {
            var data = new tbl_diagcodes();
            try
            {
                tbl_diagcodes obj = new tbl_diagcodes();
                obj.Id = id;
                data = _services.GetOne(obj);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        [AllowAnonymous]
        public IActionResult Edit(tbl_diagcodes model)
        {
            try
            {
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
                tbl_diagcodes obj = new tbl_diagcodes();
                obj.Id = id;
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
                string cnd = "  and cmp_id=" + cmpid + "  and (bodypart like '%" + searchValue + "%' or Description like '%" + searchValue + "%' or  DiagCode like '%" + searchValue + "%') and cmp_id=" + cmpid;
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_diagcodes).GetProperties()[Convert.ToInt32(sortColumn)];
                    if (sortColumnDirection.ToUpper() == "ASC")
                    {
                        Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                    }
                    else
                        Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
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

        [HttpPost]
        public IActionResult ImportData(IFormFile postedFile)
        {
            try
            {
                if (postedFile != null && postedFile.Length > 0)
                {
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
                        if (!string.IsNullOrEmpty(dt.Rows[i]["BodyPart"].ToString()))
                        {
                            tbl_diagcodes obj = new tbl_diagcodes()
                            {
                                cmp_id = cmpid,
                                BodyPart = dt.Rows[i]["BodyPart"].ToString(),
                                DiagCode = dt.Rows[i]["DiagCode"].ToString(),
                                Description = dt.Rows[i]["Description"].ToString(),
                                PreSelect = false,
                                display_order = string.IsNullOrEmpty(dt.Rows[i]["DisplayOrder"].ToString()) ? 0 : Convert.ToInt16(dt.Rows[i]["DisplayOrder"].ToString()),
                                CreatedBy = userid,
                                CreatedDate = System.DateTime.Now
                            };

                            string cnd = " and DiagCode=" + dt.Rows[i]["DiagCode"].ToString() + " and cmp_id=" + cmpid;

                            var objIsExist = _services.GetAll(cnd);

                            if (objIsExist.Count == 0)
                                _services.Insert(obj);
                            else
                            { 
                                obj.Id= objIsExist[0].Id;
                                _services.Update(obj);
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "ImportData");
            }
            return RedirectToAction("Index");
        }


        [HttpGet]
        public IActionResult ExportToExcel()
        {
            try
            {
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                string cnd = "  and cmp_id=" + cmpid;
                var data = _services.GetAll(cnd); // Retrieve all attorneys from the database

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[5]
                {
                    new DataColumn("BodyPart", typeof(string)),
                    new DataColumn("DiagCode", typeof(string)),
                    new DataColumn("Description", typeof(string)),
                    new DataColumn("PreSelect", typeof(bool)),
                    new DataColumn("DisplayOrder", typeof(int)),
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var diagcode in data)
                {
                    dt.Rows.Add(diagcode.BodyPart, diagcode.DiagCode, diagcode.Description, diagcode.PreSelect, diagcode.display_order);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "DiagCodes" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DiagCodes.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
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

        public DataTable Read2007Xlsx(IFormFile postedFile)
        {
            DataTable dt = new DataTable();
            try
            {
                if (postedFile != null && postedFile.Length > 0)
                {
                    // Read the uploaded Excel file using Open XML
                    using (Stream stream = postedFile.OpenReadStream())
                    {

                        using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(stream, false))
                        {
                            WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                            IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                            }
                            foreach (Row row in rows) //this will also include your header row...
                            {
                                DataRow tempRow = dt.NewRow();
                                int columnIndex = 0;
                                foreach (Cell cell in row.Descendants<Cell>())
                                {
                                    // Gets the column index of the cell with data
                                    int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                                   // int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellValue.InnerText));
                                    cellColumnIndex--; //zero based index
                                    if (columnIndex < cellColumnIndex)
                                    {
                                        do
                                        {
                                            tempRow[columnIndex] = ""; //Insert blank data here;
                                            columnIndex++;
                                        }
                                        while (columnIndex < cellColumnIndex);
                                    }//end if block
                                    tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);
                                    columnIndex++;
                                }//end inner foreach loop
                                dt.Rows.Add(tempRow);
                            }//end outer foreach loop
                        }//end using block
                        dt.Rows.RemoveAt(0); //...so i'm taking it out here.
                    }
                }
            }//end try
            catch (Exception ex)
            {
                SaveLog(ex, "Read2007Xlsx");
            }

            return dt;
        }//end Read2007Xlsx method

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

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return "";
            }
            string value = cell.CellValue.InnerXml;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }//end GetCellValue method

        [HttpGet]
        public ActionResult DownloadDocument()
        {
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "DiagCodes.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "DiagCodes.xlsx");
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
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(spreadsheetDocument, cell);

                        if (isFirstRow)
                        {
                            // Use the first row to add columns to the DataTable
                            dataTable.Columns.Add(cellValue);
                        }
                        else
                        {
                            // Clean HTML tags and add data to the DataTable
                            //string cleanCellValue = System.Text.RegularExpressions.Regex.Replace(cellValue, "<.*?>", string.Empty);
                            string cleanCellValue = cellValue;
                            dataRow[columnIndex] = cleanCellValue;
                        }
                        columnIndex++;
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
        #endregion
    }
}
