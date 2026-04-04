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
    public class InscosController : Controller
    {
        private readonly ILogger<InscosController> _logger;
        private readonly IMapper _mapper;
        private readonly InscosService _services = new InscosService();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;

        public InscosController(IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment, 
                                ILogger<InscosController> logger, IConfiguration configuration)
        {
            _mapper = mapper;
            _logger = logger;
            Environment = environment;
            Configuration = configuration;
        }

        public IActionResult Index()
        {           
            return View();
        }

        public IActionResult Create()
        {
            tbl_inscos obj = new tbl_inscos();
            return View(obj);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(tbl_inscos model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    model.createdby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.createddate = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    _services.Insert(model);
                }
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return RedirectToAction("Index");
        }
        public IActionResult Edit(int id)
        {
            var data = new tbl_inscos();
            try
            {
                tbl_inscos obj = new tbl_inscos();
                obj.id = id;
                data = _services.GetOne(id);
            }
            catch(Exception ex)
            {
                SaveLog(ex,"Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_inscos model)
        {
            try
            {
                model.updatedby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.updatedate = System.DateTime.Now;
                _services.Update(model);
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Delete(int id)
        {
            try
            {
                tbl_inscos obj = new tbl_inscos();
                obj.id = id;
                _services.Delete(obj);
            }
            catch(Exception ex)
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
                string cnd ="and cmp_id=" + cmpid + " and cmpname like '%" + searchValue + "%' ";
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_inscos).GetProperties()[Convert.ToInt32(sortColumn)];
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
        [HttpPost]
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
                        if (!string.IsNullOrEmpty(dt.Rows[i]["InsCo"].ToString()))
                        {
                            tbl_inscos obj = new tbl_inscos()
                            {
                                cmp_id = cmpid,
                                cmpname = (dt.Columns.Contains("InsCo") && dt.Rows[i]["InsCo"] != DBNull.Value ? dt.Rows[i]["InsCo"].ToString() : ""),
                                address1 = (dt.Columns.Contains("Address1") && dt.Rows[i]["Address1"] != DBNull.Value ? dt.Rows[i]["Address1"].ToString() : ""),
                                address2 = (dt.Columns.Contains("Address2") && dt.Rows[i]["Address2"] != DBNull.Value ? dt.Rows[i]["Address2"].ToString() : ""),
                                city = (dt.Columns.Contains("City") && dt.Rows[i]["City"] != DBNull.Value ? dt.Rows[i]["City"].ToString() : ""),
                                state = (dt.Columns.Contains("State") && dt.Rows[i]["State"] != DBNull.Value ? dt.Rows[i]["State"].ToString() : ""),
                                emailid = (dt.Columns.Contains("EmailAddress") && dt.Rows[i]["EmailAddress"] != DBNull.Value ? dt.Rows[i]["EmailAddress"].ToString() : ""),
                                telephone = (dt.Columns.Contains("Telephone") && dt.Rows[i]["Telephone"] != DBNull.Value ? dt.Rows[i]["Telephone"].ToString() : ""),
                                contactpersonname = (dt.Columns.Contains("ContactPerson") && dt.Rows[i]["ContactPerson"] != DBNull.Value ? dt.Rows[i]["ContactPerson"].ToString() : ""),
                                setasdefault = false,
                                isactive = true,
                                zipcode = (dt.Columns.Contains("Zip") && dt.Rows[i]["Zip"] != DBNull.Value ? dt.Rows[i]["Zip"].ToString() : ""),
                                faxno = "",
                                createdby = userid,
                                createddate = System.DateTime.Now
                            };

                            _services.Insert(obj);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "List");
            }
            return RedirectToAction("Index");
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

        [HttpGet]
        public ActionResult DownloadDocument()
        {
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "InsComp.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "InsComp.xlsx");
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
                dt.Columns.AddRange(new DataColumn[12]
                {
                    new DataColumn("InsCo", typeof(string)),
                    new DataColumn("Address1", typeof(string)),
                    new DataColumn("Address2", typeof(string)),
                    new DataColumn("City", typeof(string)),
                    new DataColumn("State", typeof(string)),
                    new DataColumn("EmailAddress", typeof(string)),
                    new DataColumn("Telephone", typeof(string)),
                    new DataColumn("ContactPerson", typeof(string)),
                    new DataColumn("setasdefault", typeof(bool)),
                    new DataColumn("Zip", typeof(string)),
                    new DataColumn("FaxNo",typeof(string)),
                    new DataColumn("isactive",typeof(bool))
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var inscos in data)
                {
                    dt.Rows.Add(inscos.cmpname, inscos.address1, inscos.address2, inscos.city, inscos.state, inscos.emailid, inscos.telephone,
                        inscos.contactpersonname, inscos.setasdefault, inscos.zipcode, inscos.faxno, inscos.isactive);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Inscos" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Inscos.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }

        #region private Method
        private void SaveLog(Exception ex,string acctionname)
        {
            var msg = "";
            if(ex.InnerException == null)
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
        #endregion
    }
}
