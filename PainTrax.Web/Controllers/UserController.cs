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
    public class UserController : Controller
    {
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private readonly IMapper _mapper;
        private readonly ILogger<UserController> _logger;
        private readonly UserService _services = new UserService();
        private readonly DesinationServices _designationservices = new DesinationServices();
        private readonly GroupsService _groupservices = new GroupsService();
        private readonly Common _common = new Common();
        private readonly ILoggingService _log;
        private IConfiguration _configuration;


        public UserController(IMapper mapper, ILogger<UserController> logger, ILoggingService log, IConfiguration configuration, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _mapper = mapper;
            _logger = logger;
            _log = log;
            _configuration = configuration;
            Environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Create()
        {
            tbl_users obj = new tbl_users();
            try
            {
                int? cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                ViewBag.dgList = _common.GetDesignation(cmp_id.Value);
                ViewBag.gpList = _common.GetGroups(cmp_id.Value);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return View(obj);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(tbl_users model, IFormFile? signature)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    model.createdby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.createddate = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                   // model.password = EncryptionHelper.Encrypt(model.password);
                    if (signature != null)
                    {
                        string folderPath = Path.Combine(Environment.WebRootPath, "Uploads/Sign", model.cmp_id.ToString());
                        if (!Directory.Exists(folderPath))
                        {
                            Directory.CreateDirectory(folderPath);
                        }
                        // Generate a unique filename 
                        string fileName = Guid.NewGuid().ToString() + Path.GetExtension(signature.FileName);
                        string filePath = Path.Combine(folderPath, fileName);

                        using (var fileStream = new FileStream(filePath, FileMode.Create))
                        {
                            signature.CopyTo(fileStream);
                        }
                        model.signature = fileName;
                    }
                    _services.Insert(model);
                }
                else
                {
                    var errors = ModelState.Select(x => x.Value.Errors)
                                           .Where(y => y.Count > 0)
                                           .ToList();
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
            var data = new tbl_users();
            try
            {
                tbl_users obj = new tbl_users();
                obj.Id = id;
                data = _services.GetOne(obj);
                
                data.password = EncryptionHelper.Decrypt(data.password);
                int? cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                ViewBag.dgList = _common.GetDesignation(cmp_id.Value);
                ViewBag.gpList = _common.GetGroups(cmp_id.Value);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_users model, IFormFile signature)
        {
            try
            {
                model.updatedby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.updateddate = System.DateTime.Now;
                model.password = EncryptionHelper.Encrypt(model.password);
                if (signature != null)
                {
                    string folderPath = Path.Combine(Environment.WebRootPath, "Uploads/Sign", model.cmp_id.ToString());
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    // Generate a unique filename 
                    string fileName = Guid.NewGuid().ToString() + Path.GetExtension(signature.FileName);
                    string filePath = Path.Combine(folderPath, fileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        signature.CopyTo(fileStream);
                    }
                    model.signature = fileName;
                }
                else
                {
                    model.signature = model.signature;
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
                tbl_users obj = new tbl_users();
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

                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                string cnd = " and (fname like '%" + searchValue + "%')";

                cnd = " and cmp_id=" + cmpid;

                if (!String.IsNullOrEmpty(searchValue))
                {
                    cnd += " and (fname like '%" + searchValue + "%' or lname like '%" + searchValue + "%' or emailid like '%" + searchValue + "%' or fullname like '%" + searchValue + "%' Or ";
                    cnd += " uname like '%" + searchValue + "%' or address like '%" + searchValue + "%'  )";
                }


                var designation = HttpContext.Session.GetString(SessionKeys.SessionDesignation);

                if (designation == "Provider")
                {
                    cnd += " and desig_name in ('Provider','Staff')";
                }
                else if (designation == "Client Admin")
                {
                    cnd += " and desig_name in ('Provider','Staff','Client Admin')";
                }
                else if (designation == "Admin")
                {
                    cnd += " and desig_name in ('Provider','Staff','Client Admin','Admin','FD')";
                }

                var Data = _services.GetAll(cnd);



                //sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var _sortColumn = Convert.ToInt32(sortColumn);

                    if (_sortColumn > 0)
                        _sortColumn = _sortColumn - 1;

                    var property = typeof(vm_cm_user).GetProperties()[_sortColumn];
                    //var property = s_order[Convert.ToInt32(sortColumn)];
                    if (sortColumnDirection.ToUpper() == "ASC")
                    {
                        Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                    }
                    else
                    {
                        Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                    }
                }
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
                        if (!string.IsNullOrEmpty(dt.Rows[i]["FName"].ToString()))
                        {
                            tbl_users obj = new tbl_users()
                            {
                                cmp_id = cmpid,
                                fname = dt.Rows[i]["FName"].ToString(),
                                lname = dt.Rows[i]["LName"].ToString(),
                                emailid = dt.Rows[i]["Email"].ToString(),
                                address = dt.Rows[i]["Address"].ToString(),
                                fullname = dt.Rows[i]["FullName"].ToString(),
                                uname = dt.Rows[i]["UserName"].ToString(),
                                password = dt.Rows[i]["Password"].ToString(),
                                groupid = string.IsNullOrEmpty(dt.Rows[i]["GroupId"].ToString()) ? 0 : Convert.ToInt16(dt.Rows[i]["GroupId"].ToString()),
                                desigid = string.IsNullOrEmpty(dt.Rows[i]["DesigId"].ToString()) ? 0 : Convert.ToInt16(dt.Rows[i]["DesigId"].ToString()),
                                phoneno = dt.Rows[i]["PhoneNo"].ToString(),
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
            var path = Path.Combine(Environment.WebRootPath, "Uploads/Sample", "Users.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            return File(fs, "application/octet-stream", "Users.xlsx");
        }

        [HttpGet]
        public IActionResult ExportToExcel()
        {
            try
            {
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                string cnd = " and cmp_id=" + cmpid;


                var designation = HttpContext.Session.GetString(SessionKeys.SessionDesignation);

                if (designation == "Provider")
                {
                    cnd += " and desig_name in ('Provider','Staff')";
                }
                else if (designation == "Client Admin")
                {
                    cnd += " and desig_name in ('Provider','Staff','Client Admin')";
                }
                else if (designation == "Admin")
                {
                    cnd += " and desig_name in ('Provider','Staff','Client Admin','Admin','FD')";
                }

                var data = _services.GetAll(cnd); // Retrieve all attorneys from the database

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {
                    new DataColumn("FName", typeof(string)),
                    new DataColumn("LName", typeof(string)),
                    new DataColumn("Emailaddress", typeof(string)),
                    new DataColumn("Address", typeof(string)),
                    new DataColumn("FullName", typeof(string)),
                    new DataColumn("UserName", typeof(string)),
                    new DataColumn("GroupName", typeof(string)),
                    new DataColumn("DesignationName", typeof(string)),                   
                   // new DataColumn("PhoneNo", typeof(string)),                                   
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var user in data)
                {
                    dt.Rows.Add(user.fname, user.lname, user.emailid, user.address, user.fullname, user.uname, user.group_name, user.desig_name);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Users" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Users.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }

        //public IActionResult EditUserProfile(int id)
        //{
        //    var data = new tbl_users();
        //    try
        //    {
        //        int? userId = HttpContext.Session.GetInt32(SessionKeys.SessionUserId);
        //        data = _services.GetOneById(id);
        //        int? cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

        //    }
        //    catch (Exception ex)
        //    {
        //        SaveLog(ex, "Edit");
        //    }
        //    return View(data);
        //}
        public IActionResult EditUserProfile()
        {
            var data = new tbl_users();
            try
            {
                int? userId = HttpContext.Session.GetInt32(SessionKeys.SessionUserId);
                if (userId.HasValue)
                {
                    data = _services.GetOneById(userId.Value);
                    data.password = EncryptionHelper.Decrypt(data.password);
                }
                else
                {
                    return RedirectToAction("Index", "Home");
                }
            }
            catch (Exception ex)
            {
                // Log the exception (implement SaveLog method as per your logging mechanism)
                SaveLog(ex, "EditUserProfile");
                // Handle the exception appropriately
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult EditUserProfile(tbl_users model, IFormFile signature)
        {
            try
            {
                model.updatedby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                model.updateddate = System.DateTime.Now;
                model.password = EncryptionHelper.Encrypt(model.password);
                if (signature != null)
                {
                    string folderPath = Path.Combine(Environment.WebRootPath, "Uploads/Sign", model.cmp_id.ToString());
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    // Generate a unique filename 
                    string fileName = Guid.NewGuid().ToString() + Path.GetExtension(signature.FileName);
                    string filePath = Path.Combine(folderPath, fileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        signature.CopyTo(fileStream);
                    }
                    model.signature = fileName;
                }
                else
                {
                    model.signature = model.signature_hidden;
                }
                _services.UpdateUserProfile(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "EditUserProfile");
            }
            return View("EditUserProfile", model);
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
