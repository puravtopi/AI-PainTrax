using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class MacrosController : Controller
    {
        private readonly ILogger<MacrosController> _logger;
        private readonly IMapper _mapper;
        private readonly MacrosMasterService _services = new MacrosMasterService();
        private readonly Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;

        public MacrosController(IMapper mapper, ILogger<MacrosController> logger, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _mapper = mapper;
            _logger = logger;
            Environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Create()
        {
            tbl_macros_master obj = new tbl_macros_master();
            return View(obj);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(tbl_macros_master model)
        {
            try
            {                 
                if (ModelState.IsValid)
                {
                    model.created_by = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.created_date = System.DateTime.Now;
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



        public IActionResult CreatePreOP()
        {
            tbl_macros_master obj = new tbl_macros_master();
            return View(obj);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult CreatePreOP(tbl_macros_master model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    model.created_by = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.created_date = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    _services.Insert(model);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "CreatePre");
            }
            return RedirectToAction("Index");
        }
        public IActionResult CreatePostOP()
        {
            tbl_macros_master obj = new tbl_macros_master();
            return View(obj);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult CreatePostOP(tbl_macros_master model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    model.created_by = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.created_date = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    _services.Insert(model);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "CreatePostOp");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Edit(int id)
        {
            var data = new tbl_macros_master();
            try
            {
                tbl_macros_master obj = new tbl_macros_master();
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
        public IActionResult Edit(tbl_macros_master model)
        {
            try
            {
                model.updated_by = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.updated_date = System.DateTime.Now;
                _services.Update(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return RedirectToAction("Index");
        }


        public IActionResult EditPreOP(int id)
        {
            var data = new tbl_macros_master();
            try
            {
                tbl_macros_master obj = new tbl_macros_master();
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
        public IActionResult EditPreOP(tbl_macros_master model)
        {
            try
            {
                model.updated_by = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.updated_date = System.DateTime.Now;
                _services.Update(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "EditPreOP");
            }
            return RedirectToAction("Index");
        }
        public IActionResult EditPostOP(int id)
        {
            var data = new tbl_macros_master();
            try
            {
                tbl_macros_master obj = new tbl_macros_master();
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
        public IActionResult EditPostOP(tbl_macros_master model)
        {
            try
            {
                model.updated_by = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.updated_date = System.DateTime.Now;
                _services.Update(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "EditPostOP");
            }
            return RedirectToAction("Index");
        }


        public IActionResult Delete(int id)
        {
            try
            {
                tbl_macros_master obj = new tbl_macros_master();
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
                string cnd = " and cmp_id=" + cmpid + " and (bodypart like '%" + searchValue + "%' or heading like '%" + searchValue + "%')";
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_macros_master).GetProperties()[Convert.ToInt32(sortColumn)];
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
                    DataTable dt = this.Read2007Xlsx(postedFile);
                    int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(dt.Rows[i]["Bodypart"].ToString()))
                        {
                            tbl_macros_master obj = new tbl_macros_master()
                            {
                                cmp_id = cmpid,
                                bodypart = dt.Rows[i]["Bodypart"].ToString(),
                                heading = dt.Rows[i]["Heading"].ToString(),
                                cc_desc = dt.Rows[i]["CC Description"].ToString(),
                                pe_desc = dt.Rows[i]["PE Description"].ToString(),
                                a_desc = dt.Rows[i]["Assesment/Diagnoses"].ToString(),
                                p_desc = dt.Rows[i]["P Description"].ToString(),
                                rom_desc = dt.Rows[i]["ROM Description"].ToString(),
                                cf = false,
                                pn = false,
                                pre_select = false,                               
                                created_by = userid,
                                created_date = System.DateTime.Now
                            };

                            _services.Insert(obj);
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
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "Macros.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "Macros.xlsx");
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
                dt.Columns.AddRange(new DataColumn[10]
                {
                    new DataColumn("Bodypart", typeof(string)),
                    new DataColumn("Heading", typeof(string)),
                    new DataColumn("CC Description", typeof(string)),
                    new DataColumn("PE Description", typeof(string)),
                    new DataColumn("Assesment/Diagnoses", typeof(string)),
                    new DataColumn("P Description", typeof(string)),
                    new DataColumn("ROM Description", typeof(string)),
                    new DataColumn("CF", typeof(bool)),
                    new DataColumn("PN", typeof(bool)),
                    new DataColumn("pre_select", typeof(bool)),
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var macros in data)
                {
                    dt.Rows.Add(macros.bodypart, macros.heading, macros.cc_desc, macros.pe_desc, macros.a_desc, macros.p_desc,
                        macros.rom_desc, macros.cf, macros.pn, macros.pre_select);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Macros" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Macros.xlsx");
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
        #endregion
    }
}
