using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using MySqlX.XDevAPI.Common;
using Org.BouncyCastle.Asn1.Ocsp;
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
    public class DesignationController : Controller
    {
        private readonly ILogger<tbl_designation> _logger;
        private readonly IMapper _mapper;
        private readonly DesinationServices _services = new DesinationServices();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;

        public DesignationController(IMapper mapper, ILogger<tbl_designation> logger, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment, IConfiguration configuration)
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
            tbl_designation obj = new tbl_designation();
            try
            {
                ViewBag.isError = false;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return View(obj);
        }

        [HttpPost]
        [AllowAnonymous]
        public IActionResult Create(tbl_designation model)
        {
            try
            {
                int result = 0;
                if (ModelState.IsValid)
                {
                    model.createdby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.createddate = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    result = _services.Insert(model);
                }
                if (result == -1)
                {
                    ViewBag.isError = true;
                    ViewBag.errorMsg = "Designation already present!!";

                    return View(model);
                }
                else
                {
                    ViewBag.isError = false;
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
            tbl_designation data = new tbl_designation();
            try
            {
                tbl_designation obj = new tbl_designation();
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
        public IActionResult Edit(tbl_designation model)
        {
            try
            {
                model.updatedby = 1;
                model.updateddate = System.DateTime.Now;
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
                tbl_designation obj = new tbl_designation();
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
                string cnd = cnd = " and cmp_id=" + cmpid + "  and title like '%" + searchValue + "%' "; ;
                var Data = _services.GetAll(cnd);

                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_designation).GetProperties()[Convert.ToInt32(sortColumn)];
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
            if (postedFile != null)
            {
                //Create a Folder.
                string path = Path.Combine(this.Environment.WebRootPath, "Uploads");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                // Save the uploaded Excel file.
                string filePath = Path.Combine(postedFile.FileName);
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    postedFile.CopyTo(stream);
                }

                // Read the connection string for the Excel file.
                string connectionString = Configuration.GetConnectionString("ExcelConString");

                // Modify the connection string with the file path.
                connectionString = string.Format(connectionString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(connectionString))
                {
                    connExcel.Open();
                    DataTable dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    connExcel.Close();

                    using (OleDbCommand cmdExcel = new OleDbCommand($"SELECT * FROM [{sheetName}]", connExcel))
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter(cmdExcel))
                        {
                            DataTable dt = new DataTable();
                            odaExcel.Fill(dt);

                            int? cmpId = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                            int? userId = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);

                            //foreach (DataRow row in dt.Rows)
                            //{
                            //    tbl_designation obj = new tbl_designation()
                            //    {
                            //        cmp_id = cmpId,
                            //        title = row[0].ToString(),
                            //        code = row[1].ToString(),
                            //        createdby = userId,
                            //        createddate = DateTime.Now
                            //    };

                            //    _services.Insert(obj);
                            //}
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (!string.IsNullOrEmpty(dt.Rows[i]["DesignationTitle"].ToString()))
                                {
                                    tbl_designation obj = new tbl_designation()
                                    {
                                        cmp_id = cmpId,
                                        title = dt.Rows[i]["DesignationTitle"].ToString(),
                                        code = dt.Rows[i]["Code"].ToString(),
                                        createddate = DateTime.Now,
                                        createdby = userId,

                                    };
                                    _services.Insert(obj);
                                }
                            }
                        }
                    }
                }
            }
            return RedirectToAction("Index");
        }

        [HttpGet]
        public ActionResult DownloadDocument()
        {
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "Designations.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "Designations.xlsx");
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
                dt.Columns.AddRange(new DataColumn[2]
                {
                    new DataColumn("DesignationTitle", typeof(string)),
                    new DataColumn("Code", typeof(string)),

                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var designation in data)
                {
                    dt.Rows.Add(designation.title, designation.code);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Designations" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Designations.xlsx");
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
