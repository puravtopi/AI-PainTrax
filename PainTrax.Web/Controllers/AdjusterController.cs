using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Text.RegularExpressions;


namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class AdjusterController : Controller
    {
        #region Variables
        private readonly IMapper _mapper;
        private readonly AadjusterService _services = new AadjusterService();
        private readonly ILogger<AdjusterController> _logger;
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        
        #endregion

        #region Public Method
        public AdjusterController(IMapper mapper, ILogger<AdjusterController> logger, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
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
            tbl_adjuster obj = new tbl_adjuster();
            return View(obj);
        }

        [HttpPost]
        [AllowAnonymous]
        public IActionResult Create(tbl_adjuster model)
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
                SaveLog(ex, "Create");
            }
            return RedirectToAction("Index");
        }

        public IActionResult Edit(int id)
        {
            tbl_adjuster data = new tbl_adjuster();

            try
            {
                tbl_adjuster obj = new tbl_adjuster();
                obj.id = id;
                data = _services.GetOne(id);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_adjuster model)
        {
            // model.updatedby = 1;
            //model.updateddate = System.DateTime.Now;

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
                tbl_adjuster obj = new tbl_adjuster();
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
                string cnd = " and cmp_id=" + cmpid + " and Adjuster like '%" + searchValue + "%' ";
                var Data = _services.GetAll(cnd);

                //Sorting

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var _sortColumn = Convert.ToInt32(sortColumn);

                    if (_sortColumn > 0)
                        _sortColumn = _sortColumn - 1;

                    var property = typeof(tbl_adjuster).GetProperties()[_sortColumn];
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
                    DataTable dt = this.Read2007Xlsx(postedFile);
                    int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(dt.Rows[i]["Adjuster"].ToString()))
                        {
                            tbl_adjuster obj = new tbl_adjuster()
                            {
                                cmp_id = cmpid,
                                adjuster = dt.Rows[i]["Adjuster"].ToString(),
                                address1 = dt.Rows[i]["Address1"].ToString(),
                                address2 = dt.Rows[i]["Address2"].ToString(),
                                city = dt.Rows[i]["City"].ToString(),
                                state = dt.Rows[i]["State"].ToString(),
                                zip = dt.Rows[i]["Zip"].ToString(),
                                emailaddress = dt.Rows[i]["Emailaddress"].ToString(),
                                telephone = dt.Rows[i]["Telephone"].ToString(),
                                contactperson = dt.Rows[i]["Contactperson"].ToString(),
                                fax = dt.Rows[i]["Fax"].ToString(),
                                set_as_default = false                               
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
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "Adjuster.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "Adjuster.xlsx");
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
                    new DataColumn("Adjuster", typeof(string)),
                    new DataColumn("Address1", typeof(string)),
                    new DataColumn("Address2", typeof(string)),
                    new DataColumn("City", typeof(string)),
                    new DataColumn("State", typeof(string)),
                    new DataColumn("Zip", typeof(string)),
                    new DataColumn("Emailaddress", typeof(string)),
                    new DataColumn("Telephone", typeof(string)),
                    new DataColumn("ContactPerson", typeof(string)),
                    new DataColumn("Fax", typeof(string)),
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var adjuster in data)
                {
                    dt.Rows.Add(adjuster.adjuster, adjuster.address1, adjuster.address2, adjuster.city, adjuster.state, adjuster.zip,
                        adjuster.emailaddress, adjuster.telephone, adjuster.contactperson, adjuster.fax);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Adjusters" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Adjusters.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }

        #endregion

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
            //tbl_log obj = new tbl_log();
            //obj.Message = "";
            //obj.CreatedDate = DateTime.Now;
            //obj.CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

            //logService _lg = new logService();
            //_lg.Insert(obj);

        }
        #endregion
    }
}
