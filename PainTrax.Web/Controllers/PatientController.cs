using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using MS.Models;
using MS.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.Data;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class PatientController : Controller
    {
        private readonly ILogger<PatientController> _logger;
        private readonly IMapper _mapper;
        private readonly PatientService _services = new PatientService();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private IConfiguration Configuration;

        public PatientController(IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment, 
                                 ILogger<PatientController> logger , IConfiguration configuration)
        {
            _mapper = mapper;
            _logger = logger;
            Environment = environment;
            Configuration = configuration;
        }

        public IActionResult Index(string searchtxt = "")
        {            
            return View();
        }

        public IActionResult Create()
        {
            tbl_patient obj = new tbl_patient();
            return View(obj);
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(tbl_patient model)
        {
            try
            {                
                if (ModelState.IsValid)
                {
                    model.createdby = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    model.createddate = System.DateTime.Now;
                    model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    int age = calculateAge(model.dob.Value);
                    model.age = age;
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
            var data = new tbl_patient();
            try
            {               
                tbl_patient obj = new tbl_patient();
                obj.id = id;
                data = _services.GetOne(id);
            }
            catch(Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_patient model)
        {
            try
            {
                int age = calculateAge(model.dob.Value);
                model.age = age;

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
                tbl_patient obj = new tbl_patient();
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
                string cnd = " and cmp_id=" + cmpid;
                if (!String.IsNullOrEmpty(searchValue))
                {
                    cnd += " and (fname like '%" + searchValue + "%' or lname like '%" + searchValue + "%' or CONCAT(fname,' ',lname)  LIKE '%" + searchValue + "%' or CONCAT(lname,' ',fname)  LIKE '%" + searchValue + "%' or email like '%" + searchValue + "%' or mobile like '%" + searchValue + "%' Or ";
                    cnd += " dob like '%" + searchValue + "%')";
                }
                var Data = _services.GetAll(cnd);
                //Sorting
                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_patient).GetProperties()[Convert.ToInt32(sortColumn)];
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

        private int calculateAge(DateTime bday)
        {
            DateTime today = DateTime.Today;

            int age = today.Year - bday.Year;

            if (today.Month < bday.Month ||
            ((today.Month == bday.Month) && (today.Day < bday.Day)))
            {
                age--;  //birthday in current year not yet reached, we are 1 year younger ;)
                        //+ no birthday for 29.2. guys ... sorry, just wrong date for birth
            }
            return age;
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
                        if (!string.IsNullOrEmpty(dt.Rows[i]["DOB"].ToString()))
                        {
                            DateTime dob = DateTime.FromOADate(Convert.ToDouble(dt.Rows[i]["DOB"].ToString()));
                            tbl_patient obj = new tbl_patient()
                            {
                                cmp_id = cmpid,
                                fname = dt.Rows[i]["FirstName"].ToString(),
                                lname = dt.Rows[i]["LastName"].ToString(),
                                mname = dt.Rows[i]["MiddleName"].ToString(),
                                city = dt.Rows[i]["City"].ToString(),
                                state = dt.Rows[i]["State"].ToString(),
                                email = dt.Rows[i]["eMail"].ToString(),
                                dob = dob,
                                home_ph = dt.Rows[i]["Phone"].ToString(),
                                mobile = dt.Rows[i]["Phone2"].ToString(),
                                address = dt.Rows[i]["Address1"].ToString() + " " + dt.Rows[i]["Address1"].ToString(),
                                zip = dt.Rows[i]["Zip"].ToString(),
                                gender = Common.GetGenderFromMrMrs(dt.Rows[i]["Sex"].ToString()),
                                createdby = userid,
                                createddate = System.DateTime.Now
                            };

                            obj.age = calculateAge(dob);
                            //obj.age = 1;


                            _services.Insert(obj);
                        }
                    }
                }
            }
            catch(Exception ex)
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
            var path = Path.Combine(Environment.WebRootPath + "/Uploads/Sample", "Patient.xlsx");
            var fs = new FileStream(path, FileMode.Open);

            // Return the file. A byte array can also be used instead of a stream
            return File(fs, "application/octet-stream", "Patient.xlsx");
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
                    new DataColumn("FirstName", typeof(string)),
                    new DataColumn("LastName", typeof(string)),
                    new DataColumn("MiddleName", typeof(string)),
                    new DataColumn("City", typeof(string)),
                    new DataColumn("State", typeof(string)),
                    new DataColumn("Emailaddress", typeof(string)),
                    new DataColumn("DateOfBirth", typeof(DateTime)),
                    new DataColumn("Phone", typeof(string)),
                    new DataColumn("Phone2", typeof(string)),
                    new DataColumn("Address1", typeof(string)),
                    new DataColumn("Zip", typeof(string)),
                    new DataColumn("Sex", typeof(string)),
                    new DataColumn("CreatedDate", typeof(DateTime))
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var patient in data)
                {
                    dt.Rows.Add(patient.fname, patient.lname, patient.mname, patient.city, patient.state, patient.email,
                        patient.dob, patient.home_ph, patient.mobile, patient.address, patient.zip, patient.gender, patient.createddate);
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
                    var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Patient" };
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Patient.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }

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
        #endregion
    }
}
