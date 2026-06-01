using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using static PainTrax.Web.Helper.EnumHelper;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Bibliography;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using MailKit;
using Org.BouncyCastle.Asn1.Ocsp;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Microsoft.AspNetCore.Mvc.Rendering;
using MySql.Data.MySqlClient;
using DocumentFormat.OpenXml.Office.Word;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using System.IO;
using SkiaSharp;
using System.Diagnostics.Metrics;
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;
using System.Globalization;
using MS.Models;
using MS.Services;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Net.Http;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class ReportController : Controller
    {


        private readonly ILogger<ReportController> _logger;
        private readonly POCServices _services = new POCServices();
        private readonly ProSXServices _servicesProSX = new ProSXServices();
        private readonly IVFRServices _servicesIVFR = new IVFRServices();
        private readonly ProSXDetailsServices _servicesProSXDetails = new ProSXDetailsServices();
        private readonly DailyCountServices _servicesDailyCount = new DailyCountServices();
        private readonly PtsIEServices _servicesPtsIE = new PtsIEServices();
        private readonly MDTImportServices _servicesMDTImport = new MDTImportServices();
        private readonly Common _commonservices = new Common();
        private readonly SurgeryCentreService _surgeryCentreService = new SurgeryCentreService();
        private readonly POCConfigService _pocConfigservices = new POCConfigService();
        private readonly PocStatusService _pocStatusService = new PocStatusService();
        private readonly InsuranceStatusService _insuranceStatusService = new InsuranceStatusService();
        private readonly ProBSServices _servicesProBS = new ProBSServices();
        private readonly PdfProcCodeService _servicesProcCode = new PdfProcCodeService();

        #region Patient Import Report  
        private readonly PatientImportServices _servicesPatientImport = new PatientImportServices();
        private readonly PatientService _patientservices = new PatientService();
        private readonly InscosService _inscosservices = new InscosService();
        private readonly AttorneysService _attorneyservices = new AttorneysService();
        private readonly EmpService _empService = new EmpService();
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly UserService _userService = new UserService();
        #endregion

        public ReportController(ILogger<ReportController> logger)
        {
            _logger = logger;
        }
        [HttpGet]
        public IActionResult POCReport()
        {

            var objPOC = new POCReportVM();
            objPOC.lstPOCReport = new List<POCReportVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            string cnd = " and cmp_id=" + cmpid.Value;
            var data1 = _pocConfigservices.GetAllone(cnd);
            ViewBag.Columns = data1.Select(x => x.columns).ToList();
            ViewBag.cmpid = cmpid.Value;

            objPOC._executed = false;
            objPOC._requested = false;
            objPOC._scheduled = false;



            return View(objPOC);
        }

        [HttpPost]
        public IActionResult POCReport(DateTime? fdate, DateTime? tdate, int locationid = 0, int mcodeid = 0, bool _executed = false, bool _requested = false, bool _scheduled = false)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = " where pm.cmp_id=" + cmpid.ToString();
            if (locationid > 0)
            {
                query += " and lc.id =" + locationid;
            }

            if (mcodeid == 1)
            {
                query += " and pp.inhouseprocbit=1 ";
            }
            else if (mcodeid == 2)
            {
                query += " and pp.Other =1 ";
            }
            else if (mcodeid == 3)
            {
                // query += " and pp.INhouseProcbit<>1 and  ISNULL(pp.Other,0) <> 1  ";
                query += " and IFNULL(pp.INhouseProcbit,0)<>1 and IFNULL(pp.inout,0)<>1  ";
            }
            else if (mcodeid == 4)
            {
                query += " and pp.inout=1";
            }

            string _query = "";

            if (_requested)
            {
                if (fdate != null && tdate != null)
                {
                    _query = " (tp.Requested BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";
                }
            }
            if (_executed)
            {
                ViewBag.ShowTransfer = true;
                if (fdate != null && tdate != null)
                {
                    if (!string.IsNullOrEmpty(_query))
                        _query = _query + " Or (tp.Executed BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";
                    else
                        _query = " (tp.Executed BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";

                }
            }
            if (_scheduled)
            {
                ViewBag.ShowTransfer = true;
                if (fdate != null && tdate != null)
                {
                    if (!string.IsNullOrEmpty(_query))
                        _query = _query + " Or (tp.Scheduled BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";
                    else
                        _query = " (tp.Scheduled BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";

                }
            }

            if (!string.IsNullOrEmpty(_query))
            {
                query = query + " and (" + _query + ")";
            }


            var data = _services.GetPOCReport(query);
            var objPOC = new POCReportVM();
            objPOC.lstPOCReport = data;
            TempData["query"] = query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            ViewBag.cmpid = cmpid.Value;
            string cnd = " and cmp_id=" + cmpid;
            var data1 = _pocConfigservices.GetAllone(cnd);
            ViewBag.Columns = data1.Select(x => x.columns).ToList();


            var sdata = _surgeryCentreService.GetAll(cnd + " ORDER BY Surgerycenter_name asc");
            var list = new List<SelectListItem>();



            foreach (var item in sdata)
            {
                list.Add(new SelectListItem
                {
                    Text = item.Surgerycenter_name.ToString(),
                    Value = item.Id.ToString()
                });
            }
            ViewBag.surgoryList = list;

            return View(objPOC);

        }

        [HttpPost]
        public IActionResult TransferToExecute(string ids, string dates)
        {
            ids = ids.TrimStart(',');
            dates = dates.TrimStart(',');

            if (!string.IsNullOrEmpty(ids))
            {
                var arrayId = ids.Split(',');
                var arrayDates = dates.Split(",");

                for (int i = 0; i < arrayId.Length; i++)
                {
                    if (!string.IsNullOrEmpty(arrayDates[i]))
                        _services.TransferToExecute(arrayId[i], arrayDates[i]);
                }
            }
            return Json(1);
        }

        [HttpPost]
        public IActionResult TransferToResheduled(string ids, string sDate = "")
        {
            ids = ids.TrimStart(',');

            if (!string.IsNullOrEmpty(ids))
            {
                var arrayId = ids.Split(',');

                for (int i = 0; i < arrayId.Length; i++)
                {
                    _services.TransferToReschedules(arrayId[i], sDate);
                }
            }
            return Json(1);
        }
        [HttpPost]
        public IActionResult UpdatePOCReport(string ids, string side = "", string level = "")
        {
            ids = ids.TrimStart(',');
            side = side.TrimStart(',');
            level = level.TrimStart(',');

            if (!string.IsNullOrEmpty(ids))
            {
                if (!string.IsNullOrEmpty(side) && !string.IsNullOrEmpty(level))
                {
                    var arrayId = ids.Split(',');
                    var arraySide = side.Split(',');
                    var arrayLevel = level.Split(',');

                    for (int i = 0; i < arrayId.Length; i++)
                    {
                        _services.UpdatePOCReportSideandLevel(arrayId[i], arraySide[i], arrayLevel[i]);
                    }
                }

            }
            return Json(1);
        }

        [HttpPost]
        public IActionResult UpdatePOCSideLevelReport(string id, string side = "", string level = "")
        {

            _services.UpdatePOCReportSideandLevel(id, side, level);

            return Json(1);
        }

        //public IActionResult ExportToExcel()
        //{
        //    try
        //    {
        //        string query = TempData["query"].ToString();
        //        var data = _services.GetPOCReport(query);

        //        // Create a new DataTable
        //        DataTable dt = new DataTable();
        //        // Add columns to the DataTable
        //        dt.Columns.AddRange(new DataColumn[]
        //        {
        //            new DataColumn("Sex", typeof(string)),
        //            new DataColumn("Name", typeof(string)),
        //            //new DataColumn("Provider", typeof(string)),
        //            new DataColumn("Case", typeof(string)),
        //           // new DataColumn("DOE", typeof(string)),
        //            new DataColumn("DOB", typeof(string)),
        //            new DataColumn("DOA", typeof(string)),
        //            new DataColumn("MCODE", typeof(string)),
        //            new DataColumn("Side", typeof(string)),
        //            new DataColumn("Level", typeof(string)),
        //            new DataColumn("Phone", typeof(string)),
        //            new DataColumn("Location", typeof(string)),                 
        //            new DataColumn("Ins Co", typeof(string)),
        //            new DataColumn("Claim Number", typeof(string)),
        //            //new DataColumn("Policy No", typeof(string)),
        //            new DataColumn("MC", typeof(string)),
        //            new DataColumn("Allergies", typeof(string)),
        //            new DataColumn("Requested", typeof(string)),
        //            new DataColumn("Scheduled", typeof(string)),
        //            new DataColumn("Executed", typeof(string)),
        //            new DataColumn("Note", typeof(string)),                   

        //           // new DataColumn("PhoneNo", typeof(string)),                                   
        //        });

        //        // Populate the DataTable with data from the list of attorneys
        //        foreach (var user in data)
        //        {
        //            var sex = string.IsNullOrEmpty(user.gender) ? "" : (user.gender.ToLower() == "male" ? "M" : "F");
        //            dt.Rows.Add(sex,user.name, user.providerName, user.casetype, user.doe == null ? "" : user.doe.Value.ToShortDateString(), user.dob == null ? "" : user.dob.Value.ToShortDateString(), user.doa == null ? "" : user.doa.Value.ToShortDateString(), user.mcode, user.sides, user.level, user.phone, user.location,user.cmpname, user.primary_claim_no, user.primary_policy_no, user.mc, user.allergies, user.requested == null ? "" : user.requested.Value.ToShortDateString(), user.scheduled == null ? "" : user.scheduled.Value.ToShortDateString(), user.executed == null ? "" : user.executed.Value.ToShortDateString(), user.note);
        //        }

        //        // Create a new Excel file
        //        var memoryStream = new MemoryStream();
        //        using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        //        {
        //            var workbookPart = document.AddWorkbookPart();
        //            workbookPart.Workbook = new Workbook();

        //            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //            var sheetData = new SheetData();
        //            worksheetPart.Worksheet = new Worksheet(sheetData);

        //            var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
        //            var sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Users" };
        //            sheets.Append(sheet);

        //            var (defaultStyleIndex, headerStyleIndex) = CreateStyles(workbookPart.AddNewPart<WorkbookStylesPart>());

        //            var headerRow = new Row();
        //            foreach (DataColumn column in dt.Columns)
        //            {
        //                var cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(column.ColumnName), StyleIndex = headerStyleIndex };
        //                headerRow.AppendChild(cell);
        //            }
        //            sheetData.AppendChild(headerRow);

        //            foreach (DataRow row in dt.Rows)
        //            {
        //                var newRow = new Row();
        //                foreach (var value in row.ItemArray)
        //                {
        //                    var cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(value.ToString()) };
        //                    newRow.AppendChild(cell);
        //                }
        //                sheetData.AppendChild(newRow);
        //            }
        //            // ✅ Enable AutoFilter for the header row
        //            string lastColumn = GetExcelColumnName(dt.Columns.Count);
        //            var autoFilter = new AutoFilter() { Reference = $"A1:{lastColumn}1" };
        //            worksheetPart.Worksheet.Append(autoFilter);                    

        //        }

        //        memoryStream.Seek(0, SeekOrigin.Begin);
        //        return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "POCReport.xlsx");
        //    }
        //    catch (Exception ex)
        //    {
        //        // Log or handle the exception as needed
        //        return Content("Error: " + ex.Message);
        //    }
        //}        
        public IActionResult ExportToExcel()
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cnd = " and cmp_id=" + cmpid.Value;
                var data1 = _pocConfigservices.GetAlloneExport(cnd);
                var columnList = data1.Select(x => x.columns).ToList();

                string query = TempData["query"].ToString();
                var data = _services.GetPOCReport(query);

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable

                foreach (var col in columnList)
                {
                    dt.Columns.Add(new DataColumn(col, typeof(string)));
                }

                // Populate the DataTable with data from the list of attorneys
                foreach (var user in data)
                {
                    DataRow dr = dt.NewRow();

                    foreach (var col in columnList)
                    {
                        switch (col)
                        {
                            case "Sex":
                                dr[col] = string.IsNullOrEmpty(user.gender) ? "" : (user.gender.ToLower() == "male" ? "M" : "F");
                                break;
                            case "Name":
                                dr[col] = user.name;
                                break;
                            case "Account_No":
                                dr[col] = user.acct_no;
                                break;

                            case "Case":
                                dr[col] = user.casetype;
                                break;
                            case "DOE":
                                dr[col] = user.doe?.ToShortDateString() ?? "";
                                break;
                            case "DOB":
                                dr[col] = user.dob?.ToShortDateString() ?? "";
                                break;
                            case "DOA":
                                dr[col] = user.doa?.ToShortDateString() ?? "";
                                break;
                            case "MCODE":
                                dr[col] = user.mcode;
                                break;
                            case "Phone":
                                dr[col] = user.phone;
                                break;
                            case "Location":
                                dr[col] = user.location;
                                break;
                            case "Insurance":
                                dr[col] = user.cmpname;
                                break;
                            case "Side":
                                dr[col] = user.sides;
                                break;
                            case "Level":
                                dr[col] = user.level;
                                break;
                            case "ClaimNo":
                                dr[col] = user.primary_claim_no;
                                break;
                            case "MC":
                                dr[col] = user.mc;
                                break;
                            case "MC_Details":
                                dr[col] = user.mc_details;
                                break;
                            case "Allergies":
                                dr[col] = user.allergies;
                                break;
                            case "Request":
                                dr[col] = user.requested?.ToShortDateString() ?? "";
                                break;
                            case "Scheduled":
                                dr[col] = user.scheduled?.ToShortDateString() ?? "";
                                break;
                            case "Executed":
                                dr[col] = user.executed?.ToShortDateString() ?? "";
                                break;
                            case "Note":
                                dr[col] = user.note;
                                break;
                            case "SC_Name":
                                dr[col] = user.sx_center_name;
                                break;
                            case "Surgeon":
                                dr[col] = user.surgon_name;
                                break;
                            case "Assistant":
                                dr[col] = user.assistent_name;
                                break;
                            case "Policy":
                                dr[col] = user.PolicyNo;
                                break;
                            case "Attorney":
                                dr[col] = user.Attorney;
                                break;
                            case "Vaccinated":
                                dr[col] = user.vaccinated;
                                break;
                            default:
                                dr[col] = "";
                                break;
                        }
                    }

                    dt.Rows.Add(dr);
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

                    var (defaultStyleIndex, headerStyleIndex) = CreateStyles(workbookPart.AddNewPart<WorkbookStylesPart>());

                    var headerRow = new Row();
                    foreach (DataColumn column in dt.Columns)
                    {
                        var cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(column.ColumnName), StyleIndex = headerStyleIndex };
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
                    // ✅ Enable AutoFilter for the header row
                    string lastColumn = GetExcelColumnName(dt.Columns.Count);
                    var autoFilter = new AutoFilter() { Reference = $"A1:{lastColumn}1" };
                    worksheetPart.Worksheet.Append(autoFilter);

                }

                memoryStream.Seek(0, SeekOrigin.Begin);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "POCReport.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }
        private static (uint defaultStyleIndex, uint headerStyleIndex) CreateStyles(WorkbookStylesPart stylesPart)
        {
            Stylesheet stylesheet = new Stylesheet();

            // ✅ Fonts
            DocumentFormat.OpenXml.Spreadsheet.Fonts fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                new Font(new Color { Rgb = "FF000000" }), // 0 - black font (default)
                new Font(new Color { Rgb = "FFFFFFFF" })  // 1 - white font (for header)
            );
            stylesheet.Fonts = fonts;

            // ✅ Fills
            Fills fills = new Fills();
            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None })); // 0 - none
            fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 })); // 1 - gray125 (required)
            fills.Append(new Fill( // 2 - blue fill
                new PatternFill(
                    new ForegroundColor { Rgb = "4d4dff" } // blue
                )
                { PatternType = PatternValues.Solid }
            ));
            stylesheet.Fills = fills;

            // ✅ Borders
            Borders borders = new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()); // default border
            stylesheet.Borders = borders;

            // ✅ CellFormats
            CellFormats cellFormats = new CellFormats();

            // 0 - Default (black text, no fill)
            cellFormats.Append(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 });

            // 1 - Header (white text + blue background)
            cellFormats.Append(new CellFormat
            {
                FontId = 1,
                FillId = 2,
                BorderId = 0,
                ApplyFont = true,
                ApplyFill = true
            });

            stylesheet.CellFormats = cellFormats;
            stylesPart.Stylesheet = stylesheet;
            stylesPart.Stylesheet.Save();

            return (0, 1); // (defaultStyleIndex, headerStyleIndex)
        }
        // Helper: Convert column index to Excel column letter (A, B, ..., Z, AA, AB, etc.)
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
        [HttpGet]
        public IActionResult ProSXReport()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var objPro = new ProSXReportVM();
            objPro.lstProSXReport = new List<ProSXReportVM>();

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            // ViewBag.dateList= _servicesProSX.GetProSXReportDate(cmpid.Value);
            var dates = _servicesProSX.GetProSXReportDate(cmpid.Value.ToString()); // return List<DateTime>

            var dateList = new List<SelectListItem>
{
    new SelectListItem { Value = "", Text = "--Select Date--", Selected = true }
};

            dateList.AddRange(dates.Select(d => new SelectListItem
            {
                Value = Convert.ToDateTime(d).ToString("yyyy-MM-dd"), // ✅ machine-readable
                Text = Convert.ToDateTime(d).ToString("MM/dd/yyyy")  // ✅ user-friendly
            }));

            ViewBag.dateList = dateList;

            return View(objPro);
        }

        [HttpPost]
        public IActionResult ProSXReport(DateTime? fdate, DateTime? tdate, int locationid = 0, int mcodeid = 0, bool _executed = false, bool _requested = false, bool _scheduled = false)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = " where pm.cmp_id=" + cmpid.ToString();

            string _query = "";

            if (fdate != null)
            {
                _query = " (tp.Scheduled = '" + fdate.Value.ToString("yyyy/MM/dd") + "' )";
            }
            if (tdate != null)
            {
                _query = " (tp.Scheduled = '" + tdate.Value.ToString("yyyy/MM/dd") + "' )";
            }

            if (fdate != null && tdate != null)
            {
                _query = " (tp.Scheduled = '" + fdate.Value.ToString("yyyy/MM/dd") + "' )";


            }

            if (!string.IsNullOrEmpty(_query))
            {
                query = query + " and (" + _query + ")";
            }


            var data = _servicesProSX.GetProSXReport(query);
            var objPOC = new ProSXReportVM();
            objPOC.lstProSXReport = data;
            TempData["ProSXquery"] = query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            var dates = _servicesProSX.GetProSXReportDate(cmpid.Value.ToString()); // return List<DateTime>

            var dateList = new List<SelectListItem>
{
    new SelectListItem { Value = "", Text = "--Select Date--", Selected = true }
};

            dateList.AddRange(dates.Select(d => new SelectListItem
            {
                Value = Convert.ToDateTime(d).ToString("yyyy-MM-dd"), // ✅ machine-readable
                Text = Convert.ToDateTime(d).ToString("MM/dd/yyyy")  // ✅ user-friendly
            }));

            ViewBag.dateList = dateList;
            // for status dropdown. 
            var statusList = _pocStatusService.GetAll();

            ViewBag.StatusList = new SelectList(statusList, "Name", "Name");

            // for insurance status drop down. 

            var InsurancestatusList = _insuranceStatusService.GetAll();

            ViewBag.InsurancestatusList = new SelectList(InsurancestatusList, "Name", "Name");

            string cnd = " and cmp_id=" + cmpid;

            var surgoryList = _surgeryCentreService.GetAll(cnd + " ORDER BY Surgerycenter_name asc");

            ViewBag.surgoryList = new SelectList(surgoryList, "Surgerycenter_name", "Surgerycenter_name");
            return View(objPOC);

        }

        public IActionResult ExportToExcelProSX()
        {
            try
            {
                string query = TempData["ProSXquery"].ToString();
                var data = _servicesProSX.GetProSXReport(query);

                // Build DataTable
                DataTable dt = new DataTable();
                dt.Columns.Add("SrNo");
                dt.Columns.Add("Sex");
                dt.Columns.Add("Name");
                dt.Columns.Add("Location");
                dt.Columns.Add("CaseType");
                dt.Columns.Add("MCODE");
                dt.Columns.Add("Phone");
                dt.Columns.Add("DOB");
                dt.Columns.Add("ClaimNumber");
                dt.Columns.Add("Insurance");
                dt.Columns.Add("Allergies");
                dt.Columns.Add("MC");
                dt.Columns.Add("MC_Status");
                dt.Columns.Add("Scheduled");
                dt.Columns.Add("sxCenterName");
                dt.Columns.Add("status");
                dt.Columns.Add("Color");   // color name / hex
                dt.Columns.Add("InsNote");
                dt.Columns.Add("statusInsurance");
                dt.Columns.Add("verificationComment");
                dt.Columns.Add("preopstatus");
                dt.Columns.Add("bookingSheetStatus");

                int counter = 1;

                foreach (var proSX in data)
                {
                    dt.Rows.Add(
                        counter++.ToString(),
                        proSX.gender,
                        proSX.name,
                        proSX.location,
                        proSX.casetype,
                        proSX.mcode,
                        proSX.Phone,
                        proSX.DOB,
                        proSX.ClaimNumber,
                        proSX.Insurance,
                        proSX.Allergies,
                        proSX.mc,
                        proSX.mc_Status,
                        proSX.scheduled,
                        proSX.sx_center_name,
                        proSX.sx_Status,
                        proSX.color,
                        proSX.sx_Notes,
                        proSX.SX_Ins_Ver_Status,
                        proSX.Ver_comment,
                        proSX.Preop_notesent,
                        proSX.Bookingsheet_sent
                    );
                }

                // Create Excel
                var memoryStream = new MemoryStream();

                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    //------------------------------------
                    // CREATE VALID STYLESHEET
                    //------------------------------------
                    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    var stylesheet = new Stylesheet();

                    // FONTS (required)
                    var fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
                    fonts.Append(new Font());
                    fonts.Count = 1;

                    // FILLS (Excel requires min 2)
                    var fills = new Fills();
                    fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None }));     // 0
                    fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }));   // 1
                    fills.Count = 2;

                    // BORDERS (required)
                    var borders = new Borders();
                    borders.Append(new DocumentFormat.OpenXml.Spreadsheet.Border());
                    borders.Count = 1;

                    // CELL STYLE FORMATS (required)
                    var cellStyleFormats = new CellStyleFormats();
                    cellStyleFormats.Append(new CellFormat());
                    cellStyleFormats.Count = 1;

                    // CELL FORMATS
                    var cellFormats = new CellFormats();
                    cellFormats.Append(new CellFormat());
                    cellFormats.Count = 1;

                    // attach to stylesheet
                    stylesheet.Append(fonts);
                    stylesheet.Append(fills);
                    stylesheet.Append(borders);
                    stylesheet.Append(cellStyleFormats);
                    stylesheet.Append(cellFormats);

                    stylesPart.Stylesheet = stylesheet;

                    // cache styles
                    Dictionary<string, uint> colorStyleCache = new Dictionary<string, uint>();

                    //------------------------------------
                    // COLOR NAME → HEX
                    //------------------------------------
                    string ConvertColorNameToHex(string colorName)
                    {
                        if (string.IsNullOrWhiteSpace(colorName)) return null;

                        colorName = colorName.Trim().ToLower();

                        return colorName switch
                        {
                            "red" => "FF0000",
                            "green" => "00FF00",
                            "yellow" => "FFFF00",
                            "orange" => "FFA500",
                            "blue" => "0000FF",
                            "pink" => "FFC0CB",
                            "purple" => "800080",
                            "black" => "000000",
                            "gray" => "808080",
                            _ => colorName.Replace("#", "").ToUpper()
                        };
                    }

                    //------------------------------------
                    // CREATE COLOR STYLES
                    //------------------------------------
                    uint GetStyleForColor(string colorName)
                    {
                        string hex = ConvertColorNameToHex(colorName);
                        if (hex == null) return 0;

                        if (colorStyleCache.ContainsKey(hex))
                            return colorStyleCache[hex];

                        // ADD FILL
                        fills.Append(new Fill(
                            new PatternFill(
                                new ForegroundColor() { Rgb = hex }
                            )
                            { PatternType = PatternValues.Solid }
                        ));
                        fills.Count = (uint)fills.ChildElements.Count;

                        uint fillId = (uint)(fills.Count - 1);

                        // ADD FORMAT
                        var format = new CellFormat()
                        {
                            FillId = fillId,
                            ApplyFill = true
                        };

                        cellFormats.Append(format);
                        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

                        uint styleIndex = (uint)(cellFormats.Count - 1);
                        colorStyleCache[hex] = styleIndex;

                        return styleIndex;
                    }

                    //------------------------------------
                    // SAVE STYLES EXACTLY ONCE
                    //------------------------------------
                    stylesheet.Save();

                    //------------------------------------
                    // WORKSHEET
                    //------------------------------------
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Users"
                    });


                    //------------------------------------
                    // REMOVE COLOR COLUMN FROM SHEET
                    //------------------------------------

                    // Save color values first (these match row order exactly)
                    List<string> rowColors = new List<string>();
                    foreach (DataRow row in dt.Rows)
                        rowColors.Add(row["Color"]?.ToString());

                    // Remove Color column so Excel does NOT show it
                    dt.Columns.Remove("Color");



                    // HEADER
                    var headerRow = new Row();
                    foreach (DataColumn col in dt.Columns)
                        headerRow.Append(new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(col.ColumnName)
                        });

                    sheetData.Append(headerRow);

                    // DATA
                    int rowIndex = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        var r = new Row();

                        // uint styleIndex = GetStyleForColor(row["Color"]?.ToString());
                        uint styleIndex = GetStyleForColor(rowColors[rowIndex++]);

                        foreach (var value in row.ItemArray)
                        {
                            r.Append(new Cell()
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(value?.ToString()),
                                StyleIndex = styleIndex
                            });
                        }

                        sheetData.Append(r);
                    }
                }

                memoryStream.Seek(0, SeekOrigin.Begin);
                return File(memoryStream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "ProSX Report.xlsx");
            }
            catch (Exception ex)
            {
                return Content("Error: " + ex.Message);
            }
        }

        [HttpGet]
        public IActionResult ProBSReport()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var objPro = new ProBSReportVM();
            objPro.lstProBSReport = new List<ProBSReportVM>();

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            return View(objPro);
        }

        [HttpPost]
        public IActionResult ProBSReport(DateTime? fdate)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = " where pm.cmp_id=" + cmpid.ToString();

            string _query = "";

            if (fdate != null)
            {
                _query = " (tp.Scheduled = '" + fdate.Value.ToString("yyyy/MM/dd") + "' )";
            }

            if (!string.IsNullOrEmpty(_query))
            {
                query = query + " and (" + _query + ")";
            }


            var data = _servicesProBS.GetProBSReport(query);
            var objPOC = new ProBSReportVM();
            objPOC.lstProBSReport = data;
            TempData["ProBSquery"] = query;

            //ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            //// for status dropdown. 
            //var statusList = _pocStatusService.GetAll();

            //ViewBag.StatusList = new SelectList(statusList, "Name", "Name");



            string cnd = " and cmp_id=" + cmpid;
            return View(objPOC);

        }

        public IActionResult ExportToExcelProBS()
        {
            try
            {
                string query = TempData["ProBSquery"].ToString();
                var data = _servicesProBS.GetProBSReport(query);

                // Build DataTable
                DataTable dt = new DataTable();
                dt.Columns.Add("SrNo");
                dt.Columns.Add("Sex");
                dt.Columns.Add("Name");
                dt.Columns.Add("Account_No");
                dt.Columns.Add("DOB");
                dt.Columns.Add("Phone");
                dt.Columns.Add("Address");
                dt.Columns.Add("City");
                dt.Columns.Add("State");
                dt.Columns.Add("Zip");
                dt.Columns.Add("SSN");
                dt.Columns.Add("Case");
                dt.Columns.Add("DOE");
                dt.Columns.Add("DOA");
                dt.Columns.Add("AttorneyName");
                dt.Columns.Add("AttorneyPhone");
                dt.Columns.Add("Mcode");
                dt.Columns.Add("Sides");   // color name / hex
                dt.Columns.Add("Level");
                dt.Columns.Add("location");
                dt.Columns.Add("Insurance");
                dt.Columns.Add("ClaimNumber");
                dt.Columns.Add("WCB");
                dt.Columns.Add("scheduled");
                dt.Columns.Add("sx_center_name");
                dt.Columns.Add("Booking Status");
                dt.Columns.Add("Note");
                dt.Columns.Add("Procedures");
                dt.Columns.Add("cptcodes");
                dt.Columns.Add("icdcodes");
                dt.Columns.Add("specialequ");
                int counter = 1;
                List<string> rowColors = new List<string>();
                Dictionary<string, tbl_pdfproccode?> mcodeCache = new Dictionary<string, tbl_pdfproccode?>();

                foreach (var proBS in data)
                {
                    // -----------------------------
                    // GET PROC DETAILS BY MCODE
                    // -----------------------------
                    tbl_pdfproccode? procData = null;
                    string cleanMcode = CleanMcode(proBS.mcode);

                    if (!string.IsNullOrEmpty(cleanMcode))
                    {
                        if (!mcodeCache.ContainsKey(cleanMcode))
                        {
                            mcodeCache[cleanMcode] = _servicesProcCode.GetMcode(cleanMcode);
                        }

                        procData = mcodeCache[cleanMcode];
                    }
                    string rowColor = null;
                    if (!string.IsNullOrEmpty(proBS.Bookingsheet_sent))
                    {
                        switch (proBS.Bookingsheet_sent.Trim().ToLower())
                        {
                            case "yes":
                                rowColor = "green";
                                break;

                            case "no":
                                rowColor = "red";
                                break;

                            case "n/a":
                            case "na":
                                rowColor = "yellow";
                                break;
                        }
                    }

                    rowColors.Add(rowColor);
                    dt.Rows.Add(
                        counter++.ToString(),
                        proBS.gender,
                        proBS.name,
                        proBS.account_no,
                        proBS.DOB?.ToShortDateString() ?? "",
                        proBS.Phone,
                        proBS.Address,
                        proBS.City,
                        proBS.State,
                        proBS.Zip,
                        proBS.SSN,
                        proBS.casetype,
                        proBS.doe?.ToShortDateString() ?? "",
                        proBS.doa?.ToShortDateString() ?? "",
                        proBS.AttorneyName,
                        proBS.AttorneyPhone,
                        proBS.mcode,
                        proBS.sides,
                        proBS.level,
                        proBS.location,
                        proBS.Insurance,
                        proBS.ClaimNumber,
                        proBS.WCB,
                        proBS.scheduled,
                        proBS.sx_center_name,
                        proBS.Bookingsheet_sent,
                        proBS.Note,

                        // 🔽 NEW COLUMNS FROM tbl_pdfproccode
                        procData?.mprocedure,
                        procData?.cptcodes,
                        procData?.icdcodes,
                        procData?.specialequ
                    );
                }

                // Create Excel
                var memoryStream = new MemoryStream();

                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    //------------------------------------
                    // CREATE VALID STYLESHEET
                    //------------------------------------
                    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    var stylesheet = new Stylesheet();

                    // FONTS (required)
                    var fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
                    fonts.Append(new Font());
                    fonts.Count = 1;

                    // FILLS (Excel requires min 2)
                    var fills = new Fills();
                    fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None }));     // 0
                    fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }));   // 1
                    fills.Count = 2;

                    // BORDERS (required)
                    var borders = new Borders();
                    borders.Append(new DocumentFormat.OpenXml.Spreadsheet.Border());
                    borders.Count = 1;

                    // CELL STYLE FORMATS (required)
                    var cellStyleFormats = new CellStyleFormats();
                    cellStyleFormats.Append(new CellFormat());
                    cellStyleFormats.Count = 1;

                    // CELL FORMATS
                    var cellFormats = new CellFormats();
                    cellFormats.Append(new CellFormat());
                    cellFormats.Count = 1;

                    // attach to stylesheet
                    stylesheet.Append(fonts);
                    stylesheet.Append(fills);
                    stylesheet.Append(borders);
                    stylesheet.Append(cellStyleFormats);
                    stylesheet.Append(cellFormats);

                    stylesPart.Stylesheet = stylesheet;

                    // cache styles
                    Dictionary<string, uint> colorStyleCache = new Dictionary<string, uint>();

                    //------------------------------------
                    // COLOR NAME → HEX
                    //------------------------------------
                    string ConvertColorNameToHex(string colorName)
                    {
                        if (string.IsNullOrWhiteSpace(colorName)) return null;

                        colorName = colorName.Trim().ToLower();

                        return colorName switch
                        {
                            "red" => "FF0000",
                            "green" => "00FF00",
                            "yellow" => "FFFF00",
                            "orange" => "FFA500",
                            "blue" => "0000FF",
                            "pink" => "FFC0CB",
                            "purple" => "800080",
                            "black" => "000000",
                            "gray" => "808080",
                            _ => colorName.Replace("#", "").ToUpper()
                        };
                    }

                    //------------------------------------
                    // CREATE COLOR STYLES
                    //------------------------------------
                    uint GetStyleForColor(string colorName)
                    {
                        string hex = ConvertColorNameToHex(colorName);
                        if (hex == null) return 0;

                        if (colorStyleCache.ContainsKey(hex))
                            return colorStyleCache[hex];

                        // ADD FILL
                        fills.Append(new Fill(
                            new PatternFill(
                                new ForegroundColor() { Rgb = hex }
                            )
                            { PatternType = PatternValues.Solid }
                        ));
                        fills.Count = (uint)fills.ChildElements.Count;

                        uint fillId = (uint)(fills.Count - 1);

                        // ADD FORMAT
                        var format = new CellFormat()
                        {
                            FillId = fillId,
                            ApplyFill = true
                        };

                        cellFormats.Append(format);
                        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

                        uint styleIndex = (uint)(cellFormats.Count - 1);
                        colorStyleCache[hex] = styleIndex;

                        return styleIndex;
                    }

                    //------------------------------------
                    // SAVE STYLES EXACTLY ONCE
                    //------------------------------------
                    stylesheet.Save();

                    //------------------------------------
                    // WORKSHEET
                    //------------------------------------
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Users"
                    });


                    ////------------------------------------
                    //// REMOVE COLOR COLUMN FROM SHEET
                    ////------------------------------------

                    //// Save color values first (these match row order exactly)
                    //List<string> rowColors = new List<string>();
                    //foreach (DataRow row in dt.Rows)
                    //    rowColors.Add(row["Color"]?.ToString());

                    //// Remove Color column so Excel does NOT show it
                    //dt.Columns.Remove("Color");



                    // HEADER
                    var headerRow = new Row();
                    foreach (DataColumn col in dt.Columns)
                        headerRow.Append(new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(col.ColumnName)
                        });

                    sheetData.Append(headerRow);

                    // DATA
                    int rowIndex = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        var r = new Row();

                        // uint styleIndex = GetStyleForColor(row["Color"]?.ToString());
                        uint styleIndex = GetStyleForColor(rowColors[rowIndex++]);

                        foreach (var value in row.ItemArray)
                        {
                            r.Append(new Cell()
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(value?.ToString()),
                                StyleIndex = styleIndex
                            });
                        }

                        sheetData.Append(r);
                    }
                }

                memoryStream.Seek(0, SeekOrigin.Begin);
                return File(memoryStream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "ProBS Report.xlsx");
            }
            catch (Exception ex)
            {
                return Content("Error: " + ex.Message);
            }
        }

        private string CleanMcode(string mcode)
        {
            if (string.IsNullOrWhiteSpace(mcode))
                return mcode;

            // Take value before #
            return mcode.Split('#')[0].Trim();
        }
        private Stylesheet CreateStylesheet()
        {
            return new Stylesheet(
                new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                    new Font(), // 0: default
                    new Font(new DocumentFormat.OpenXml.Spreadsheet.Bold()), // 1: bold
                    new Font(new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = 11 }, new Color { Rgb = "FF0000" }) // 2: red
                ),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "D3D3D3" }) { PatternType = PatternValues.Solid }) // 2: gray
                ),
                new Borders(new DocumentFormat.OpenXml.Spreadsheet.Border()),
                new CellFormats(
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }, // 0: default
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 0, ApplyFont = true, ApplyFill = true }, // 1: header
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }, // 2: normal
                    new CellFormat { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true } // 3: cancel
                )
            );
        }

        private DocumentFormat.OpenXml.Spreadsheet.Columns CreateColumnWidths(int count)
        {
            var columns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
            for (uint i = 1; i <= count; i++)
            {
                columns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column { Min = i, Max = i, Width = 20, CustomWidth = true });
            }
            return columns;
        }

        private SheetViews CreateFreezePane()
        {
            return new SheetViews(new SheetView
            {
                WorkbookViewId = 0,
                Pane = new Pane
                {
                    VerticalSplit = 2,
                    //TopRow = 2,
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen
                }
            });
        }

        private MergeCells CreateMergedHeaders()
        {
            return new MergeCells(
                new MergeCell { Reference = new StringValue("A1:B1") },
                new MergeCell { Reference = new StringValue("C1:K1") },
                new MergeCell { Reference = new StringValue("L1:O1") },
                new MergeCell { Reference = new StringValue("P1:T1") }
            );
        }

        private Row CreateGroupHeaderRow()
        {
            var row = new Row();
            row.Append(CreateTextCell("A1", "", 2));
            row.Append(CreateTextCell("C1", "Patient Details", 2));
            row.Append(CreateTextCell("L1", "Surgery Details", 2));
            row.Append(CreateTextCell("P1", "Insurance Verification Details", 2));
            return row;
        }

        private Row CreateHeaderRow(string[] headers)
        {
            var row = new Row();
            for (int i = 0; i < headers.Length; i++)
            {
                row.Append(CreateTextCell(GetColumnLetter(i + 1) + "2", headers[i], 2));
            }
            return row;
        }

        private Row CreateDataRow(string[] values, int rowIndex, uint styleIndex)
        {
            var row = new Row();
            for (int i = 0; i < values.Length; i++)
            {
                row.Append(CreateTextCell(GetColumnLetter(i + 1) + rowIndex.ToString(), values[i], styleIndex));
            }
            return row;
        }

        private Cell CreateTextCell(string cellReference, string cellValue, uint styleIndex)
        {
            return new Cell
            {
                CellReference = cellReference,
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue ?? ""),
                StyleIndex = styleIndex
            };
        }

        private string GetColumnLetter(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                columnIndex = (columnIndex - modulo) / 26;
            }
            return columnName;
        }


        [HttpPost]
        public ActionResult UpdateProSXReport(ProSXReportVM model)
        {


            //sx_Notes = @sx_Notes,

            //  _servicesProSX.updae(model);

            try
            {
                _servicesProSX.Update(model);
            }
            catch (Exception ex)
            {
                // SaveLog(ex, "Edit");
            }
            //return RedirectToAction("Index");



            TempData["Message"] = "Details saved successfully!";
            return RedirectToAction("ProSXReport");
        }

        [HttpGet]
        public IActionResult IVFRReport()
        {

            var objPro = new IVFRReportVM();
            objPro.lstIVFRReport = new List<IVFRReportVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            //objPro._executed = false;
            //objPro._requested = false;
            objPro._scheduled = false;

            return View(objPro);
        }

        [HttpPost]
        public IActionResult IVFRReport(DateTime? fdate, DateTime? tdate, int locationid = 0, int mcodeid = 0, bool _executed = false, bool _requested = false, bool _scheduled = false)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = " where pm.cmp_id=" + cmpid.ToString();

            string _query = "";

            //if (fdate != null)
            //{
            //    _query = " (tp.Scheduled >= '" + fdate.Value.ToString("yyyy/MM/dd") + "' )";
            //}
            if (locationid > 0)
            {
                query += " and lc.id =" + locationid;
            }

            if (_scheduled)
            {
                if (fdate != null && tdate != null)
                {
                    if (!string.IsNullOrEmpty(_query))
                        _query = _query + " Or (tp.Scheduled BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";
                    else
                        _query = " (tp.Scheduled BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";

                }
            }

            if (!string.IsNullOrEmpty(_query))
            {
                query = query + " and (" + _query + ")";
            }


            var data = _servicesIVFR.GetIVFRReport(query);
            var objPOC = new IVFRReportVM();
            objPOC.lstIVFRReport = data;
            TempData["IVFRquery"] = query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View(objPOC);

        }

        public IActionResult ExportToExcelIVFR()
        {
            try
            {
                string query = TempData["IVFRquery"].ToString();
                var data = _servicesIVFR.GetIVFRReport(query);

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {
                    new DataColumn("Name", typeof(string)),
                     new DataColumn("Sex", typeof(string)),
                    new DataColumn("MCODE", typeof(string)),
                    new DataColumn("Case", typeof(string)),
                    new DataColumn("DOB", typeof(string)),
                    new DataColumn("DOA", typeof(string)),
                    new DataColumn("SSN", typeof(string)),
                    new DataColumn("Phone", typeof(string)),
                    new DataColumn("Location", typeof(string)),
                    new DataColumn("Address", typeof(string)),
                    new DataColumn("Ins Co", typeof(string)),
                    new DataColumn("Claim Number", typeof(string)),
                    new DataColumn("Policy No", typeof(string)),
                    new DataColumn("Scheduled", typeof(string)),

                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var IVFR in data)
                {
                    dt.Rows.Add(IVFR.name, IVFR.gender, IVFR.mcode, IVFR.casetype, IVFR.dob == null ? "" : IVFR.dob.Value.ToShortDateString(), IVFR.doa == null ? "" : IVFR.doa.Value.ToShortDateString(), IVFR.ssn, IVFR.phone, IVFR.location, IVFR.Address, IVFR.InsCo, IVFR.primary_claim_no, IVFR.primary_policy_no, IVFR.scheduled == null ? "" : IVFR.scheduled.Value.ToShortDateString());
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "IVFR Report.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }


        [HttpGet]
        public IActionResult DailyCountReport()
        {

            var objPro = new DailyCountReportVM();
            objPro.lstDailyCountReport = new List<DailyCountReportVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            //objPro._executed = false;
            //objPro._requested = false;
            //objPro._scheduled = false;

            return View(objPro);
        }
        [HttpPost]
        public IActionResult DailyCountReport(DateTime? fdate, DateTime? tdate)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = "";
            if (fdate != null && tdate != null)
            {
                query = query + "  '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "'";
            }

            var data = _servicesDailyCount.GetDailyCountReport(query, cmpid.ToString());
            var objPOC = new DailyCountReportVM();
            objPOC.lstDailyCountReport = data;
            TempData["DailyCountquery"] = query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View(objPOC);

        }

        public IActionResult ExportToExcelDailyCount()
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string query = TempData["DailyCountquery"].ToString();
                var data = _servicesDailyCount.GetDailyCountReport(query, cmpid.ToString());

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {
                    new DataColumn("DOE", typeof(string)),
                    new DataColumn("Location", typeof(string)),
                    new DataColumn("WC", typeof(Int64)),
                    new DataColumn("NF", typeof(Int64)),
                    new DataColumn("LIEN", typeof(Int64)),
                    new DataColumn("NoOfIE", typeof(Int64)),
                    new DataColumn("NoOfFU", typeof(Int64)),


                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var cnt in data)
                {
                    dt.Rows.Add(cnt.doe?.ToShortDateString() ?? "", cnt.location, cnt.WC, cnt.NF, cnt.LIEN, cnt.NoOFIE, cnt.NoOFFU);
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

                    //foreach (DataRow row in dt.Rows)
                    //{
                    //    var newRow = new Row();
                    //    foreach (var value in row.ItemArray)
                    //    {
                    //        var cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(value.ToString()) };
                    //        newRow.AppendChild(cell);
                    //    }
                    //    sheetData.AppendChild(newRow);
                    //}
                    foreach (DataRow row in dt.Rows)
                    {
                        var newRow = new Row();

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            var value = row[i];
                            Cell cell;

                            // Numeric columns
                            if (dt.Columns[i].ColumnName == "WC" ||
                                dt.Columns[i].ColumnName == "NF" ||
                                dt.Columns[i].ColumnName == "LIEN" ||
                                dt.Columns[i].ColumnName == "NoOfIE" ||
                                dt.Columns[i].ColumnName == "NoOfFU")
                            {
                                cell = new Cell
                                {
                                    CellValue = new CellValue(value.ToString())
                                    // IMPORTANT: no DataType for numbers
                                };
                            }
                            else // Text columns
                            {
                                cell = new Cell
                                {
                                    DataType = CellValues.String,
                                    CellValue = new CellValue(value.ToString())
                                };
                            }

                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }

                memoryStream.Seek(0, SeekOrigin.Begin);
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Daily Count Report.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }


        [HttpGet]
        public IActionResult MDTImportReport()
        {

            var objPro = new MDTImportReportVM();
            objPro.lstMDTImportReport = new List<MDTImportReportVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            //objPro._executed = false;
            //objPro._requested = false;
            //objPro._scheduled = false;

            return View(objPro);
        }
        [HttpPost]
        public IActionResult MDTImportReport(DateTime? fdate, DateTime? tdate)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = " pm.cmp_id=" + cmpid.ToString();

            string _query = "";

            if (fdate != null && tdate != null)
            {


                _query = "'" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "' and " + query;

            }

            var data = _servicesMDTImport.GeMDTImportReport(_query);
            var objPOC = new MDTImportReportVM();
            objPOC.lstMDTImportReport = data;
            TempData["MDTImportquery"] = _query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View(objPOC);

        }

        public IActionResult ExportToMDTImport()
        {
            try
            {
                string query = TempData["MDTImportquery"].ToString();
                var data = _servicesMDTImport.GeMDTImportReport(query);

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {
                    new DataColumn("DOE", typeof(string)),
                    new DataColumn("PatientIE_ID", typeof(string)),
                    new DataColumn("lname", typeof(string)),
                    new DataColumn("fname", typeof(string)),

                    new DataColumn("mname", typeof(string)),
                    new DataColumn("gender", typeof(string)),
                    new DataColumn("dob", typeof(string)),
                    new DataColumn("address", typeof(string)),
                    new DataColumn("city", typeof(string)),
                    new DataColumn("state", typeof(string)),
                    new DataColumn("zip", typeof(string)),
                    new DataColumn("home_ph", typeof(string)),

                    new DataColumn("mobile", typeof(string)),
                    new DataColumn("location", typeof(string))

                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var cnt in data)
                {
                    dt.Rows.Add(cnt.doe, cnt.PatientIE_ID, cnt.lname, cnt.fname, cnt.mname, cnt.gender, cnt.dob, cnt.address, cnt.city, cnt.state, cnt.zip, cnt.home_ph, cnt.mobile, cnt.location);
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MDT Import Report.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }




        [HttpGet]
        public IActionResult PtsIEReport()
        {

            var objPro = new PtsIEReportVM();
            objPro.lstPtsIEReport = new List<PtsIEReportVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            //objPro._executed = false;
            //objPro._requested = false;
            //objPro._scheduled = false;

            return View(objPro);
        }
        [HttpPost]
        public IActionResult PtsIEReport(DateTime? fdate, DateTime? tdate)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string _query = "";


            if (fdate != null && tdate != null)
            {


                _query = "'" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "' and ie.cmp_id = " + cmpid.ToString();

            }


            var data = _servicesPtsIE.GetPtsIEReport(_query);
            var objPOC = new PtsIEReportVM();
            objPOC.lstPtsIEReport = data;
            TempData["lstPtsIEReport"] = _query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View(objPOC);

        }

        public IActionResult ExportToExcelPtsIE()
        {
            try
            {
                string query = TempData["lstPtsIEReport"].ToString();
                var data = _servicesPtsIE.GetPtsIEReport(query);

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {

                    new DataColumn("PName", typeof(string)),
                    new DataColumn("mobile", typeof(string)),
                    new DataColumn("location", typeof(string)),
                    new DataColumn("CaseType", typeof(string)),
                    new DataColumn("doe", typeof(string)),
                    new DataColumn("doa", typeof(string)),
                    new DataColumn("Ins", typeof(string)),
                    new DataColumn("primary_policy_no", typeof(string)),
                    new DataColumn("Attorney", typeof(string)),
                    new DataColumn("LastVisit", typeof(string))


                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var cnt in data)
                {
                    dt.Rows.Add(cnt.PName, cnt.mobile, cnt.location, cnt.CaseType, cnt.doe, cnt.doa, cnt.Ins, cnt.primary_policy_no, cnt.Attorney, cnt.LastVisit);
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Pts IE Report.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }




        [HttpGet]
        public IActionResult ProSXDetailsReport()
        {

            var objPro = new ProSXDetailsReportVM();
            objPro.lstProSXDetailsReport = new List<ProSXDetailsReportVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            //objPro._executed = false;
            //objPro._requested = false;
            //objPro._scheduled = false;

            return View(objPro);
        }

        [HttpPost]
        public IActionResult ProSXDetailsReport(DateTime? fdate, DateTime? tdate, int locationid = 0)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            string query = " where ie.cmp_id=" + cmpid.ToString();

            string _query = "";


            if (locationid > 0)
            {
                query += " and lc.id =" + locationid;
            }


            if (fdate != null && tdate != null)
            {
                _query = _query + " (pd.Requested BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "') Or (pd.Scheduled BETWEEN '" + fdate.Value.ToString("yyyy/MM/dd") + "' and '" + tdate.Value.ToString("yyyy/MM/dd") + "')";

            }


            if (!string.IsNullOrEmpty(_query))
            {
                query = query + " and (" + _query + ")";
            }


            var data = _servicesProSXDetails.GetPtsIEReport(query);
            var objPOC = new ProSXDetailsReportVM();
            objPOC.lstProSXDetailsReport = data;
            TempData["ProSXDetailsReportRquery"] = query;

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View(objPOC);

        }

        public IActionResult ExportToExcelProSXDetails()
        {
            try
            {
                string query = TempData["ProSXDetailsReportRquery"].ToString();
                var data = _servicesProSXDetails.GetPtsIEReport(query);

                // Create a new DataTable
                DataTable dt = new DataTable();
                // Add columns to the DataTable
                dt.Columns.AddRange(new DataColumn[]
                {
                     new DataColumn("name", typeof(string)),
                     new DataColumn("sex", typeof(string)),
                     new DataColumn("MC", typeof(string)),
                     new DataColumn("CaseType", typeof(string)),
                     new DataColumn("location", typeof(string)),
                     new DataColumn("Vaccinated", typeof(string)),
                     new DataColumn("MCODE", typeof(string)),
                     new DataColumn("BodyPart", typeof(string)),
                     new DataColumn("Ins_ver_status", typeof(string)),
                     new DataColumn("MC_Status", typeof(string)),
                     new DataColumn("Case_Status", typeof(string)),
                     new DataColumn("InsVerStatus", typeof(string)),
                     new DataColumn("Vac_Status", typeof(string)),
                     new DataColumn("Scheduled", typeof(string)),
                     new DataColumn("Executed", typeof(string)),
                     new DataColumn("Requested", typeof(string)),
                     new DataColumn("DOE", typeof(string)),
                     new DataColumn("PolicyNo", typeof(string)),
                     new DataColumn("Attorney", typeof(string))
                });

                // Populate the DataTable with data from the list of attorneys
                foreach (var t in data)
                {
                    dt.Rows.Add(t.name, t.sex, t.mc, t.casetype, t.location, t.vaccinated, t.mcode, t.bodypart, t.ins_ver_status, t.mc_status, t.case_status, t.insverstatus, t.vac_status, t.scheduled, t.executed, t.requested, t.DOE == null ? "" : t.DOE.Value.ToShortDateString(), t.PolicyNo, t.Attorney);//, IVFR.scheduled == null ? "" : IVFR.scheduled.Value.ToShortDateString());
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
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ProSXDetails Report.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed
                return Content("Error: " + ex.Message);
            }
        }


        #region Patient Import Report   

        [HttpGet]
        public IActionResult PatientImportReport()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;

            // Nothing to load from DB on first visit — the user uploads the file.
            return View(new PatientImportReportVM());
        }

        // 1. Set up the TextInfo object to easily handle the Title Case conversion
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        // 2. Helper function to apply the casing rule (First letter capital, rest small)
        // We call .ToLower() first because ToTitleCase ignores strings that are ALL CAPS.
        string FormatNamePart(string namePart)
        {
            if (string.IsNullOrWhiteSpace(namePart)) return string.Empty;
            return textInfo.ToTitleCase(namePart.Trim().ToLower());
        }
        private int calculateAge(DateTime bday, DateTime? dos)
        {
            DateTime today = dos == null ? DateTime.Today : dos.Value;

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
        [ValidateAntiForgeryToken]
        public IActionResult PatientImportReport(
                                                    string parsedJson,
                                                    IFormFile excelFile,
                                                    string locationId,     // ← new
                                                    string locationName,
                                                    string providerId,     // ← new
                                                    string providerName,
                                                    string fdate
                                                )   // ← new (optional, for logging/display)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            int providerIDrec = Convert.ToInt32(providerId);

            // ── Re-populate ViewBag in case we need to return the view with errors ──
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            // ── Validate location ──
            if (string.IsNullOrWhiteSpace(locationId) || !int.TryParse(locationId, out int locId))
            {
                TempData["ImportError"] = "Please select a location before importing.";
                return RedirectToAction("PatientImportReport");
            }

            if (string.IsNullOrWhiteSpace(parsedJson))
            {
                TempData["ImportError"] = "No data received. Please upload and preview the file first.";
                return RedirectToAction("PatientImportReport");
            }

            try
            {
                var rows = JsonConvert.DeserializeObject<List<PatientImportRowVM>>(parsedJson);

                if (rows == null || rows.Count == 0)
                {
                    TempData["ImportError"] = "The uploaded file contained no valid rows.";
                    return RedirectToAction("PatientImportReport");
                }

                // 3. Map the data
                List<PatientImportReportVM> lstPatientImportReport = rows.Select(row =>
                {
                    // --- Parse Name ---
                    string lastName = string.Empty;
                    string firstName = null;

                    if (!string.IsNullOrWhiteSpace(row.PatientName))
                    {
                        var nameParts = row.PatientName.Split(',');

                        if (nameParts.Length > 0)
                            lastName = FormatNamePart(nameParts[0]);

                        if (nameParts.Length > 1)
                            firstName = FormatNamePart(nameParts[1]);
                    }

                    // --- Parse Dates ---
                    DateTime? DOEmain = DateTime.TryParse(fdate, out DateTime parsedDoemain) ? parsedDoemain : (DateTime?)null;

                    DateTime? dob = DateTime.TryParse(row.DOB, out DateTime parsedDob) ? parsedDob : (DateTime?)null;
                    DateTime? doa = DateTime.TryParse(row.DateOfAccident, out DateTime parsedDoa) ? parsedDoa : (DateTime?)null;
                    DateTime? DOE = DateTime.TryParse(row.DOE, out DateTime parsedDoe) ? parsedDoe : (DateTime?)null;
                    // --- Create and Map Object ---
                    return new PatientImportReportVM
                    {
                        // If last_name is entirely empty but required by the DB, provide a fallback to prevent DB validation errors
                        last_name = string.IsNullOrEmpty(lastName) ? "Unknown" : lastName,
                        first_name = firstName,

                        dob = dob,
                        doa = doa,

                        sex = row.Sex,
                        address = row.Address,
                        phone = row.Phone,
                        ssn = row.SocialSecurityNo,
                        employer_company = row.EmployerCompany,
                        employer_address = row.EmployerAddress,
                        emergency_contact = row.EmergencyName,
                        work_phone = row.WorkPhone,
                        condition_related_to = row.CaseType, // Mapped CaseType to condition_related_to based on context
                        insurance_company = row.InsuranceCompany,
                        insurance_address = row.InsAddress,
                        insurance_phone = row.InsPhone,
                        claim_number = row.ClaimNo,
                        claim_address = row.ClaimAddress,
                        nf2 = row.NF2,
                        policy_number = row.PolicyNo,
                        policy_holder = row.PolicyHolder,
                        wcb_number = row.WCBNo,
                        carrier_case_number = row.CarrierCaseNo,
                        policy_adjuster = row.PolicyAdjuster,
                        attorney = row.Attorney,
                        firm_name = row.FirmName,
                        attorney_address = row.AttorneyAddress,
                        attorney_phone = row.AttorneyPhone,
                        attorney_fax = row.AttorneyFax,
                        imported_at = DateTime.Now,
                        // If maindoe is not null, use it; otherwise fall back to DOE
                        DOE = DOE ?? DOEmain,
                        // DOE =  DOE,
                        loc_id = Convert.ToInt32(locationId)
                    };
                }).ToList();
                PatientImportReportVM model = new PatientImportReportVM
                {
                    cmpy_id = cmpid,
                    loc_id = Convert.ToInt16(locationId),
                    lstPatientImportReport = lstPatientImportReport
                };
                int result = PatientImportDataToDB(model, providerIDrec);

                if (result > 0)
                {
                    TempData["ImportSuccess"] = $"Successfully imported {lstPatientImportReport.Count()} patient record(s) to location \"{Convert.ToString(locationName)}\".";
                }
                //TempData["ImportSuccess"] = $"Successfully imported {inserted} patient record(s) to location "{ locationName}".";
                //TempData["ImportSuccess"] = "Success";// $"Successfully imported {inserted} patient record(s) to location "{ locationName}".";

            }
            catch (Exception ex)
            {
                TempData["ImportError"] = $"Import failed: {ex.Message}";
            }

            return RedirectToAction("PatientImportReport");
        }


        public int PatientImportDataToDB(PatientImportReportVM model, int providerID)
        {
            try
            {
                var data = model;
                int patientId = 0, priminsId = 0, secinsId = 0, attornyId = 0, adjusterId = 0, empId = 0;
                DateTime dos = DateTime.UtcNow.Date;
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                int age = 0;
                int ie = 0;
                //int defaultlocation = HttpContext.Session.GetInt32(SessionKeys.SessionLocationId).Value;
                foreach (var report in model.lstPatientImportReport)
                {

                    age = calculateAge(Convert.ToDateTime(report.dob).Date, dos);


                    tbl_patient objPatient = new tbl_patient()
                    {
                        //account_no = string.Empty,
                        address = report.address,
                        //city = string.Empty,
                        dob = Convert.ToDateTime(report.dob).Date,
                        email = string.Empty,
                        fname = report.first_name,
                        gender = report.sex == "Female" ? "2" : "1",
                        //home_ph = string.Empty,
                        lname = report.last_name,
                        //mc = ,
                        //mc_details = model.mc_details,
                        //mname = model.mname,
                        mobile = report.phone,
                        //handeness = model.handeness,
                        ssn = report.ssn,
                        // state = model.state,
                        // physicianid = model.physicianid,
                        // vaccinated = model.vaccinated,
                        // zip = model.zip,
                        cmp_id = cmpid,
                        createdby = userid,
                        age = age
                    };

                    patientId = _patientservices.Insert(objPatient);

                    HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, patientId);
                    string pid = HttpContext.Session.GetInt32(SessionKeys.SessionPatientId).ToString();
                    ViewBag.patientId = patientId;
                    var query = "";
                    List<tbl_inscos> insdata = new List<tbl_inscos>();
                    tbl_inscos objInscos = new tbl_inscos();

                    if (!string.IsNullOrEmpty(report.insurance_company))
                    {
                        query = " and cmpname='" + report.insurance_company + "' and cmp_id=" + cmpid + "";

                        insdata = _inscosservices.GetAll(query);

                        //save primary insurance

                        objInscos = new tbl_inscos()
                        {
                            address1 = report.insurance_address,
                            cmpname = report.insurance_company,
                            telephone = report.insurance_phone,
                            cmp_id = cmpid,
                            createdby = userid

                        };

                        if (insdata.Count > 0)
                        {
                            objInscos.id = insdata[0].id.Value;
                            _inscosservices.Update(objInscos);
                            priminsId = insdata[0].id.Value;
                        }
                        else
                        {
                            priminsId = _inscosservices.Insert(objInscos);
                        }

                    }


                    if (!string.IsNullOrEmpty(report.attorney))
                    {

                        query = " and Attorney='" + report.attorney + "' and cmp_id=" + cmpid + "";

                        var attrydata = _attorneyservices.GetAll(query);


                        //save attorney

                        tbl_attorneys objAttorneys = new tbl_attorneys()
                        {
                            Attorney = report.attorney,
                            // EmailId = model.attory_email,
                            ContactNo = report.attorney_phone,
                            cmp_id = cmpid,
                            CreatedBy = userid,
                            // Paralegal = model.paralegal


                        };

                        if (attrydata.Count > 0)
                        {
                            objAttorneys.Id = attrydata[0].Id.Value;
                            _attorneyservices.Update(objAttorneys);
                            attornyId = attrydata[0].Id.Value;
                        }
                        else
                        {
                            attornyId = _attorneyservices.Insert(objAttorneys);
                        }
                    }

                    //if (!string.IsNullOrEmpty(model.adj_name))
                    //{
                    //    query = " and adjuster='" + model.adj_name + "' and cmp_id=" + cmpid + "";

                    //    var adjdata = _aadjusterService.GetAll(query);

                    //    //save adjuster

                    //    tbl_adjuster objAdjuster = new tbl_adjuster()
                    //    {
                    //        adjuster = model.adj_name,
                    //        emailaddress = model.adj_email,
                    //        telephone = model.adj_phone,
                    //        fax = model.adj_fax_no,
                    //        cmp_id = cmpid,
                    //        created_by = userid
                    //    };

                    //    if (adjdata.Count > 0)
                    //    {
                    //        objAdjuster.id = adjdata[0].id.Value;
                    //        _aadjusterService.Update(objAdjuster);
                    //        adjusterId = adjdata[0].id.Value;
                    //    }
                    //    else
                    //    {
                    //        adjusterId = _aadjusterService.Insert(objAdjuster);
                    //    }
                    //}

                    if (!string.IsNullOrEmpty(report.employer_company))
                    {
                        query = " and name='" + report.employer_company + "' and patient_id=" + patientId + "";

                        var empdata = _empService.GetAll(query);

                        //save employee

                        tbl_emp objEmp = new tbl_emp()
                        {
                            address = report.employer_company,
                            name = report.employer_company,
                            phone = report.work_phone,
                            //fax = model.emp_fax_no,
                            patient_id = patientId
                        };

                        if (empdata.Count > 0)
                        {
                            objEmp.id = empdata[0].id.Value;
                            _empService.Update(objEmp);
                            empId = empdata[0].id.Value;
                        }
                        else
                        {
                            empId = _empService.Insert(objEmp);
                        }
                    }

                    //save IE details 

                    tbl_patient_ie objIE = new tbl_patient_ie()
                    {
                        adjuster_id = adjusterId,
                        attorney_id = attornyId,
                        created_by = userid,
                        doa = report.doa,
                        doe = Convert.ToDateTime(report.DOE),
                        emp_id = empId,
                        is_active = true,
                        location_id = report.loc_id,
                        provider_id = providerID,
                        patient_id = patientId,
                        primary_claim_no = report.claim_number,
                        primary_ins_cmp_id = priminsId,
                        primary_policy_no = report.policy_number,
                        //primary_wcb_group = model.prime_WCB_group,
                        // secondary_claim_no = model.sec_claim_no,
                        //  secondary_ins_cmp_id = secinsId,
                        // secondary_policy_no = model.sec_policy_no,
                        // secondary_wcb_group = model.sec_WCB_group,
                        compensation = report.condition_related_to,
                        // accident_type = model.accidentType,
                        // state = model.state,
                        // physicianid = model.physicianid,
                        //  alert_note = model.alert_note,
                        //  referring_physician = model.referring_physician
                    };

                    ie = _ieService.Insert(objIE);

                    // HttpContext.Session.SetInt32(SessionKeys.SessionIEId, ie);
                }
                // return Json(new { status = 1, patintid = patientId, ieid = ie });
                return ie;
                // TempData["Message"] = "Details saved successfully!";
                //  return RedirectToAction("PatientImportReport");
            }
            catch (Exception ex)
            {
                // return Json(new { status = 0 });
                return 0;
            }

        }

        #endregion

        #region Patient Import Report  pdf

        [HttpGet]
        public IActionResult PatientImportFromPDF()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;

            // Nothing to load from DB on first visit — the user uploads the file.
            return View(new PatientImportReportVM());
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult PatientImportFromPDF(
                                            string parsedJson,
                                            string locationId,     // ← new
                                            string locationName,
                                            string providerId,     // ← new
                                            string providerName,
                                            string fdate
                                        )   // ← new (optional, for logging/display)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            int providerIDrec = Convert.ToInt32(providerId);

            // ── Re-populate ViewBag in case we need to return the view with errors ──
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            // ── Validate location ──
            if (string.IsNullOrWhiteSpace(locationId) || !int.TryParse(locationId, out int locId))
            {
                TempData["ImportError"] = "Please select a location before importing.";
                return RedirectToAction("PatientImportFromPDF");
            }

            if (string.IsNullOrWhiteSpace(parsedJson))
            {
                TempData["ImportError"] = "No data received. Please upload and preview the file first.";
                return RedirectToAction("PatientImportFromPDF");
            }

            try
            {
                var rows = JsonConvert.DeserializeObject<List<PatientImportRowVM>>(parsedJson);

                if (rows == null || rows.Count == 0)
                {
                    TempData["ImportError"] = "The uploaded file contained no valid rows.";
                    return RedirectToAction("PatientImportFromPDF");
                }

                // 3. Map the data
                List<PatientImportReportVM> lstPatientImportReport = rows.Select(row =>
                {
                    // --- Parse Name ---
                    string lastName = string.Empty;
                    string firstName = null;

                    if (!string.IsNullOrWhiteSpace(row.PatientName))
                    {
                        var nameParts = row.PatientName.Split(',');

                        if (nameParts.Length > 0)
                            lastName = FormatNamePart(nameParts[0]);

                        if (nameParts.Length > 1)
                            firstName = FormatNamePart(nameParts[1]);
                    }

                    // --- Parse Dates ---
                    DateTime? DOEmain = DateTime.TryParse(fdate, out DateTime parsedDoemain) ? parsedDoemain : (DateTime?)null;

                    DateTime? dob = DateTime.TryParse(row.DOB, out DateTime parsedDob) ? parsedDob : (DateTime?)null;
                    DateTime? doa = DateTime.TryParse(row.DateOfAccident, out DateTime parsedDoa) ? parsedDoa : (DateTime?)null;
                    DateTime? DOE = DateTime.TryParse(row.DOE, out DateTime parsedDoe) ? parsedDoe : (DateTime?)null;
                    // --- Create and Map Object ---
                    return new PatientImportReportVM
                    {
                        // If last_name is entirely empty but required by the DB, provide a fallback to prevent DB validation errors
                        last_name = string.IsNullOrEmpty(lastName) ? "Unknown" : lastName,
                        first_name = firstName,

                        dob = dob,
                        doa = doa,

                        sex = row.Sex,
                        address = row.Address,
                        phone = row.Phone,
                        ssn = row.SocialSecurityNo,
                        employer_company = row.EmployerCompany,
                        employer_address = row.EmployerAddress,
                        emergency_contact = row.EmergencyName,
                        work_phone = row.WorkPhone,
                        condition_related_to = row.CaseType, // Mapped CaseType to condition_related_to based on context
                        insurance_company = row.InsuranceCompany,
                        insurance_address = row.InsAddress,
                        insurance_phone = row.InsPhone,
                        claim_number = row.ClaimNo,
                        claim_address = row.ClaimAddress,
                        nf2 = row.NF2,
                        policy_number = row.PolicyNo,
                        policy_holder = row.PolicyHolder,
                        wcb_number = row.WCBNo,
                        carrier_case_number = row.CarrierCaseNo,
                        policy_adjuster = row.PolicyAdjuster,
                        attorney = row.Attorney,
                        firm_name = row.FirmName,
                        attorney_address = row.AttorneyAddress,
                        attorney_phone = row.AttorneyPhone,
                        attorney_fax = row.AttorneyFax,
                        imported_at = DateTime.Now,
                        // If maindoe is not null, use it; otherwise fall back to DOE
                        DOE = DOE ?? DOEmain,
                        // DOE =  DOE,
                        loc_id = Convert.ToInt32(locationId)
                    };
                }).ToList();
                PatientImportReportVM model = new PatientImportReportVM
                {
                    cmpy_id = cmpid,
                    loc_id = Convert.ToInt16(locationId),
                    lstPatientImportReport = lstPatientImportReport
                };
                int result = PatientImportDataToDB(model, providerIDrec);
                //string locationName = "";
                if (result > 0)
                {
                    TempData["ImportSuccess"] = $"Successfully imported {lstPatientImportReport.Count()} patient record(s) to location \"{Convert.ToString(locationName)}\".";
                }
                //TempData["ImportSuccess"] = $"Successfully imported {inserted} patient record(s) to location "{ locationName}".";
                //TempData["ImportSuccess"] = "Success";// $"Successfully imported {inserted} patient record(s) to location "{ locationName}".";

            }
            catch (Exception ex)
            {
                TempData["ImportError"] = $"Import failed: {ex.Message}";
            }

            return RedirectToAction("PatientImportFromPDF");
        }

        #endregion

        [HttpGet]
        public IActionResult ProcedureStatusReport()
        {
            DateTime? fdate = null;
            DateTime? tdate = null;
            int locationid = 0;


            //objPro.lstSurgoryCenterDashboardVM = new List<SurgoryCenterDashboardVM>();
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            DateTime fDate = fdate == null ? new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1) : Convert.ToDateTime(fdate);
            DateTime tDate = tdate == null ? fDate.AddMonths(1).AddDays(-1) : Convert.ToDateTime(tdate);


            var Data = _services.GetSurgoryDashboardData(fDate.ToString("yyyy-MM-dd"), tDate.ToString("yyyy-MM-dd"), cmpid.ToString());
            var objPOC = new SurgoryCenterDashboardVM
            {
                fdate = fDate,
                tdate = tDate,
                lstSurgoryCenterDashboardVM = Data
            };

            string cnd = " and cmp_id=" + cmpid;
            var data = _surgeryCentreService.GetAll(cnd);
            var list = new List<SelectListItem>();



            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.Surgerycenter_name.ToString(),
                    Value = item.Id.ToString()
                });
            }
            ViewBag.surgoryList = list;

            return View(objPOC); // 
        }

        [HttpPost]
        public IActionResult ProcedureStatusReport(DateTime? fdate, DateTime? tdate, int locationid = 0)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var data = _services.GetSurgoryDashboardData(fdate.Value.ToString("yyyy-MM-dd"), tdate.Value.ToString("yyyy-MM-dd"), cmpid.ToString());


            //var data = _servicesProSXDetails.GetPtsIEReport(query);

            var objPOC = new SurgoryCenterDashboardVM
            {
                fdate = fdate,
                tdate = tdate,
                lstSurgoryCenterDashboardVM = data
            };

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            return View(objPOC);

        }

        #region Patient Import Report  pdf AI

        [HttpGet]
        public IActionResult PatientImportFromPDFAI()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            var providers = _userService.GetProviders(cmpid.Value);
            ViewBag.providerList = providers;

            ////ViewBag.ClaudeApiKey = _configuration["Claude:ApiKey"];
            //var config = HttpContext.RequestServices.GetRequiredService<IConfiguration>();
            //ViewBag.ClaudeApiKey = config["Claude:ApiKey"];

            // Nothing to load from DB on first visit — the user uploads the file.
            return View(new PatientImportReportVM());
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ExtractPdfPages([FromBody] PdfExtractRequest request)
        {
            if (request?.Pages == null || request.Pages.Count == 0)
                return BadRequest(new { error = "No page images provided." });

            // ── Read API key securely from appsettings.json ──────────────────────────

            var config = HttpContext.RequestServices.GetRequiredService<IConfiguration>();
            var apiKey = config["Claude:ApiKey"];

            if (string.IsNullOrWhiteSpace(apiKey))
                return StatusCode(500, new { error = "Claude API key is not configured on the server." });

            const string model = "claude-opus-4-5"; // best for handwritten / scanned forms

            const string prompt = @"You are a medical data extraction assistant. The images below are ALL pages of a single patient intake form (it may be handwritten, typed, or scanned across multiple pages).

Read ALL pages together and extract the patient's information into ONE single JSON object.
Do NOT create one record per page — the entire PDF is for ONE patient.

Return ONLY a valid JSON object — no markdown, no explanation, no backticks — with exactly these keys:
{
  ""Patient Name"": ""Lastname, Firstname"",
  ""DOB"": ""YYYY-MM-DD or empty string"",
  ""Sex"": ""Male or Female or Other or empty string"",
  ""Address"": ""full address including city state zip"",
  ""Phone"": """",
  ""Social Security #"": """",
  ""Employer/Company"": """",
  ""Employer Address"": """",
  ""Emergency Name"": """",
  ""Work Phone"": """",
  ""Date of Accident"": ""YYYY-MM-DD or empty string"",
  ""CaseType"": ""NF or WC or Other or empty string"",
  ""Insurance Company"": """",
  ""Ins Address"": """",
  ""Ins Phone"": """",
  ""Claim #"": """",
  ""Claim Address"": """",
  ""NF-2"": ""Yes or No or empty string"",
  ""Policy #"": """",
  ""Policy Holder"": """",
  ""WCB #"": """",
  ""Carrier Case #"": """",
  ""Policy Adjuster"": """",
  ""Attorney"": """",
  ""Firm Name"": """",
  ""Attorney Address"": """",
  ""Attorney Phone"": """",
  ""Attorney Fax"": """",
  ""DOE"": """"
}

Rules:
- ONE JSON object total — not an array, not one per page.
- Patient Name MUST be ""Lastname, Firstname"". If only one name is found, use it as lastname.
- All dates must be YYYY-MM-DD. Leave empty string if unclear.
- CaseType: ""NF"" for No-Fault/Auto, ""WC"" for Workers Comp, ""Other"" for everything else.
- If a field is not found on any page, use empty string """".
- Never include extra keys. Never add comments or explanation.";

            // ── Build the Claude request body ────────────────────────────────────────
            var contentBlocks = new List<object>();

            // Add each page as an image block
            foreach (var pageBase64 in request.Pages)
            {
                contentBlocks.Add(new
                {
                    type = "image",
                    source = new
                    {
                        type = "base64",
                        media_type = "image/png",
                        data = pageBase64
                    }
                });
            }

            // Add the extraction prompt last
            contentBlocks.Add(new { type = "text", text = prompt });

            var claudeRequest = new
            {
                model = model,
                max_tokens = 1500,
                messages = new[]
                {
            new { role = "user", content = contentBlocks }
        }
            };

            var json = System.Text.Json.JsonSerializer.Serialize(claudeRequest);

            // ── Call Anthropic API server-side ───────────────────────────────────────

            var httpClientFactory = HttpContext.RequestServices.GetRequiredService<IHttpClientFactory>();
            var httpClient = httpClientFactory.CreateClient();
            httpClient.Timeout = TimeSpan.FromMinutes(3);

            using var httpRequest = new HttpRequestMessage(HttpMethod.Post,
                "https://api.anthropic.com/v1/messages");

            httpRequest.Headers.Add("x-api-key", apiKey);
            httpRequest.Headers.Add("anthropic-version", "2023-06-01");
            httpRequest.Content = new StringContent(json, Encoding.UTF8, "application/json");

            HttpResponseMessage httpResponse;
            try
            {
                httpResponse = await httpClient.SendAsync(httpRequest);
            }
            catch (Exception ex)
            {
                return StatusCode(502, new { error = $"Could not reach Anthropic API: {ex.Message}" });
            }

            var responseBody = await httpResponse.Content.ReadAsStringAsync();

            if (!httpResponse.IsSuccessStatusCode)
            {
                // Parse Anthropic error message and return it (without leaking the API key)
                string claudeError;
                try
                {
                    using var errDoc = JsonDocument.Parse(responseBody);
                    claudeError = errDoc.RootElement
                        .GetProperty("error")
                        .GetProperty("message")
                        .GetString() ?? responseBody;
                }
                catch
                {
                    claudeError = $"HTTP {(int)httpResponse.StatusCode}";
                }

                return StatusCode((int)httpResponse.StatusCode, new { error = claudeError });
            }

            // ── Parse the Claude response ────────────────────────────────────────────
            string extractedText;
            try
            {
                using var doc = JsonDocument.Parse(responseBody);
                var contentArray = doc.RootElement.GetProperty("content");
                extractedText = string.Empty;

                foreach (var block in contentArray.EnumerateArray())
                {
                    if (block.GetProperty("type").GetString() == "text")
                    {
                        extractedText = block.GetProperty("text").GetString() ?? string.Empty;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Failed to parse Claude response: {ex.Message}" });
            }

            // ── Strip any accidental markdown fences and parse JSON ─────────────────
            var clean = extractedText
                .Trim()
                .TrimStart('`')
                .TrimEnd('`');

            // Remove ```json or ``` prefix/suffix if present
            if (clean.StartsWith("json", StringComparison.OrdinalIgnoreCase))
                clean = clean[4..].Trim();

            try
            {
                // Validate it's proper JSON, then return it as-is to the browser
                using var resultDoc = JsonDocument.Parse(clean);
                return Content(clean, "application/json");
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Claude returned invalid JSON: {ex.Message}. Raw: {clean[..Math.Min(200, clean.Length)]}" });
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult PatientImportFromPDFAI(
                                            string parsedJson,
                                            string locationId,     // ← new
                                            string locationName,
                                            string providerId,     // ← new
                                            string providerName,
                                            string fdate
                                        )   // ← new (optional, for logging/display)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            int providerIDrec = Convert.ToInt32(providerId);

            // ── Re-populate ViewBag in case we need to return the view with errors ──
            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            // ── Validate location ──
            if (string.IsNullOrWhiteSpace(locationId) || !int.TryParse(locationId, out int locId))
            {
                TempData["ImportError"] = "Please select a location before importing.";
                return RedirectToAction("PatientImportFromPDF");
            }

            if (string.IsNullOrWhiteSpace(parsedJson))
            {
                TempData["ImportError"] = "No data received. Please upload and preview the file first.";
                return RedirectToAction("PatientImportFromPDF");
            }

            try
            {
                var rows = JsonConvert.DeserializeObject<List<PatientImportRowVM>>(parsedJson);

                if (rows == null || rows.Count == 0)
                {
                    TempData["ImportError"] = "The uploaded file contained no valid rows.";
                    return RedirectToAction("PatientImportFromPDF");
                }

                // 3. Map the data
                List<PatientImportReportVM> lstPatientImportReport = rows.Select(row =>
                {
                    // --- Parse Name ---
                    string lastName = string.Empty;
                    string firstName = null;

                    if (!string.IsNullOrWhiteSpace(row.PatientName))
                    {
                        var nameParts = row.PatientName.Split(',');

                        if (nameParts.Length > 0)
                            lastName = FormatNamePart(nameParts[0]);

                        if (nameParts.Length > 1)
                            firstName = FormatNamePart(nameParts[1]);
                    }

                    // --- Parse Dates ---
                    DateTime? DOEmain = DateTime.TryParse(fdate, out DateTime parsedDoemain) ? parsedDoemain : (DateTime?)null;

                    DateTime? dob = DateTime.TryParse(row.DOB, out DateTime parsedDob) ? parsedDob : (DateTime?)null;
                    DateTime? doa = DateTime.TryParse(row.DateOfAccident, out DateTime parsedDoa) ? parsedDoa : (DateTime?)null;
                    DateTime? DOE = DateTime.TryParse(row.DOE, out DateTime parsedDoe) ? parsedDoe : (DateTime?)null;
                    // --- Create and Map Object ---
                    return new PatientImportReportVM
                    {
                        // If last_name is entirely empty but required by the DB, provide a fallback to prevent DB validation errors
                        last_name = string.IsNullOrEmpty(lastName) ? "Unknown" : lastName,
                        first_name = firstName,

                        dob = dob,
                        doa = doa,

                        sex = row.Sex,
                        address = row.Address,
                        phone = row.Phone,
                        ssn = row.SocialSecurityNo,
                        employer_company = row.EmployerCompany,
                        employer_address = row.EmployerAddress,
                        emergency_contact = row.EmergencyName,
                        work_phone = row.WorkPhone,
                        condition_related_to = row.CaseType, // Mapped CaseType to condition_related_to based on context
                        insurance_company = row.InsuranceCompany,
                        insurance_address = row.InsAddress,
                        insurance_phone = row.InsPhone,
                        claim_number = row.ClaimNo,
                        claim_address = row.ClaimAddress,
                        nf2 = row.NF2,
                        policy_number = row.PolicyNo,
                        policy_holder = row.PolicyHolder,
                        wcb_number = row.WCBNo,
                        carrier_case_number = row.CarrierCaseNo,
                        policy_adjuster = row.PolicyAdjuster,
                        attorney = row.Attorney,
                        firm_name = row.FirmName,
                        attorney_address = row.AttorneyAddress,
                        attorney_phone = row.AttorneyPhone,
                        attorney_fax = row.AttorneyFax,
                        imported_at = DateTime.Now,
                        // If maindoe is not null, use it; otherwise fall back to DOE
                        DOE = DOE ?? DOEmain,
                        // DOE =  DOE,
                        loc_id = Convert.ToInt32(locationId)
                    };
                }).ToList();
                PatientImportReportVM model = new PatientImportReportVM
                {
                    cmpy_id = cmpid,
                    loc_id = Convert.ToInt16(locationId),
                    lstPatientImportReport = lstPatientImportReport
                };
                int result = PatientImportDataToDB(model, providerIDrec);
                //string locationName = "";
                if (result > 0)
                {
                    TempData["ImportSuccess"] = $"Successfully imported {lstPatientImportReport.Count()} patient record(s) to location \"{Convert.ToString(locationName)}\".";
                }
                //TempData["ImportSuccess"] = $"Successfully imported {inserted} patient record(s) to location "{ locationName}".";
                //TempData["ImportSuccess"] = "Success";// $"Successfully imported {inserted} patient record(s) to location "{ locationName}".";

            }
            catch (Exception ex)
            {
                TempData["ImportError"] = $"Import failed: {ex.Message}";
            }

            return RedirectToAction("PatientImportFromPDFAI");
        }

        #endregion



        #region File conversion from pdf to report docx in any format.


        /// <summary>
        /// GET: /ReportGeneration/ReportGenerationFromPdf
        /// Display the main report generation page
        /// </summary>
        public IActionResult ReportGenerationFromPdf()
        {
            // string? cmpid = HttpContext.Session.GetString(SessionKeys.SessionCmpName);
            ViewBag.LocationTemplates = new List<SelectListItem>
{
    new SelectListItem { Value = "Midtown_LH.docx", Text = "Midtown_LH" },
    new SelectListItem { Value = "RBSC_LH.docx", Text = "RBSC_LH" },
    new SelectListItem { Value = "RSCR_LH.docx", Text = "RSCR_LH" },
    new SelectListItem { Value = "SCJC_LH.docx", Text = "SCJC_LH" },
    new SelectListItem { Value = "Carewll_LH.docx", Text = "Carewll_LH" },
    new SelectListItem { Value = "FASC_LH.docx", Text = "FASC_LH" },
    new SelectListItem { Value = "FASCExt_LH.docx", Text = "FASCExt_LH" }
};


            return View();
        }

        /// <summary>
        /// POST: /ReportGeneration/ExtractPdfData
        /// Accepts PDF pages (base64 PNG images), calls Claude API for extraction
        /// Returns extracted form data as JSON
        /// </summary>
        [HttpPost("extract")]

        public async Task<IActionResult> ExtractPdfData([FromBody] PdfExtractRequest request)
        {
            if (request?.Pages == null || request.Pages.Count == 0)
                return BadRequest(new { error = "No page images provided." });

            try
            {

                var config = HttpContext.RequestServices.GetRequiredService<IConfiguration>();
                var apiKey = config["Claude:ApiKey"];
                //var apiKey = _configuration["Claude:ApiKey"];
                if (string.IsNullOrWhiteSpace(apiKey))
                    return StatusCode(500, new { error = "Claude API key is not configured on the server." });

                var extractedData = await CallClaudeApiForExtraction(apiKey, request.Pages);
                return Ok(extractedData);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"PDF extraction failed: {ex.Message}" });
            }
        }

        /// <summary>
        /// POST: /ReportGeneration/GenerateReportDocx
        /// Takes extracted data and generates professional DOCX report
        /// </summary>
        [HttpPost("generate")]

        public IActionResult GenerateReportDocx([FromBody] ExtractedFormDataDto data)
        {
            if (data == null)
                return BadRequest(new { error = "No data provided." });

            try
            {
                if (string.IsNullOrWhiteSpace(data.LastName))
                    data.LastName = "Unknown";
                if (string.IsNullOrWhiteSpace(data.FirstName))
                    data.FirstName = "Patient";

                byte[] docxBytes = GenerateDocxReport(data);
                string safeFilename = SanitizeFilename(
                    $"Report_{data.LastName}_{data.FirstName}_{data.DOS?.Replace("/", "") ?? DateTime.Now.ToString("yyyyMMdd")}.docx");
                // ── Create Safe Filename ────────────────────────────────────


                return File(
                    docxBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    safeFilename);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"DOCX generation failed: {ex.Message}" });
            }
        }

        // ════════════════════════════════════════════════════════════════════════════════
        // PRIVATE HELPER METHODS
        // ════════════════════════════════════════════════════════════════════════════════

        private async Task<ExtractedFormDataDto> CallClaudeApiForExtraction(string apiKey, List<string> pageBase64List)
        {
            const string model = "claude-opus-4-6";

            const string prompt = @"You are a medical data extraction assistant.
The document is a surgical intake form (SHOULDER or KNEE — Right or Left side).
Read the form carefully. Return ONE JSON object only — no array, no markdown, no backticks.

════════════════════════════════════════════════════════
FIELD-BY-FIELD EXTRACTION RULES
════════════════════════════════════════════════════════


1. PATIENT NAME  (top-right header of the form)
   The name is printed as:  LastName, FirstName MiddleInitial
   Example: ""Barzola, Isaiah O"" → LastName=Barzola  FirstName=Isaiah O
   RULE: Everything BEFORE the comma = LastName.  Everything AFTER the comma = FirstName.

2. DOB / DOS  (top-right header)
   Format: MM/DD/YYYY. Leave empty if not found.

3. MRN  (top-right header, labelled MRN:)

4. Assistant  (the ""Assistant:"" line on the form)
   Look for a checkmark, slash, underline next to: IA. / GD.
   Only one will be marked. Return that value exactly.

5. Anesthesia  (the ""Anesthesia:"" line on the form)
   Look for a checkmark, slash, underline next to: Sed. / Block / Gen.
   Only one will be marked. Return that value exactly.

6. LocationLH  (the ""Location:"" field, handwritten or typed)
   Return the written location text exactly as appears.

7. CaseType  (the ""Case Type:"" field)
   Return the case type as written.

8. ReportTemplate  (the VERY FIRST line of the form)
   Combine body part + side: SHOULDER + RIGHT = ""RSAS"", SHOULDER + LEFT = ""LSAS"",
   KNEE + RIGHT = ""RKAS"", KNEE + LEFT = ""LKAS""

9. CPTRows (CPT Codes section)
   Same marking rule (checkmark, slash, or underline = MARKED).
   
   For lines with MULTIPLE sub-options on the same line (like ""hyperemic / hypertrophic / post-traumatic""):
   - Only include the sub-options that are individually marked/checked
   - Do NOT include unmarked sub-options even if they appear on a marked row
   
   For rows with nested/indented checked items (like supraspinatus, anterior labrum):
   - Include the parent CPT code + description (shortened, without bold label words like ""limited""/""extensive"")
   - Append each checked nested item after the parent description
   
   For bold descriptive labels that are NOT options (e.g., ""limited"", ""extensive"", ""complete""):
   - Exclude these structural/header words unless they are the only description
   
   Return as comma-separated values with format:
   [CPT Code] [Base Description],[checked sub-items joined by space]

10. ICD10Rows  (ICD-10 Codes section)
    Same marking rule. Return as comma-separated values and Look for a checkmark, slash, underline next to it on same line which is MARKED and the bold one with the text of full line.

11. IntraOpRows  (Pre-op findings)
    Same marking rule. Return as comma-separated integers.

12. PreOp  (the ""Pre-op:"" line on the form)
    Look for a checkmark, slash, underline next to: RC tear / Labrum tear / Chondral lesion / Synovitis / Adhesions / Bursitis.
    Multiple one will be marked. Return that value exactly.

13. PreOpPassiveRoM  (the ""Pre-op Passive ROM:"" line on the form and handwritten or typed)
    Look for a handwritten or typed next to: FF: / ER: / IRSS. and for Instability:. chekmark, slash on Yes,No.
    Multiple one will be marked. Return that value exactly.

════════════════════════════════════════════════════════
RETURN ONLY THIS JSON OBJECT:
════════════════════════════════════════════════════════
{
  ""FirstName"": """",
  ""LastName"": """",
  ""DOB"": """",
  ""DOS"": """",
  ""MRN"": """",
  ""Assistant"": """",
  ""Anesthesia"": """",
  ""LocationLH"": """",
  ""CaseType"": """",
  ""ReportTemplate"": """",
  ""CPTRows"": """",
  ""ICD10Rows"": """",
  ""IntraOpRows"": """",
  ""PreOp"": """",
  ""PreOpPassiveRoM"": """"

}";

            var contentBlocks = new List<object>();

            foreach (var pageBase64 in pageBase64List)
            {
                contentBlocks.Add(new
                {
                    type = "image",
                    source = new
                    {
                        type = "base64",
                        media_type = "image/png",
                        data = pageBase64
                    }
                });
            }

            contentBlocks.Add(new
            {
                type = "text",
                text = prompt
            });

            var claudeRequest = new
            {
                model = model,
                max_tokens = 2000,
                messages = new[]
                {
                    new { role = "user", content = contentBlocks }
                }
            };

            var json = System.Text.Json.JsonSerializer.Serialize(claudeRequest);


            // ── Call Anthropic API ──────────────────────────────────────────
            var httpClientFactory = HttpContext.RequestServices.GetRequiredService<IHttpClientFactory>();
            var httpClient = httpClientFactory.CreateClient();
            httpClient.Timeout = TimeSpan.FromMinutes(3);


            using var httpRequest = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
            httpRequest.Headers.Add("x-api-key", apiKey);
            httpRequest.Headers.Add("anthropic-version", "2023-06-01");
            httpRequest.Content = new StringContent(json, Encoding.UTF8, "application/json");

            HttpResponseMessage httpResponse;
            try
            {
                httpResponse = await httpClient.SendAsync(httpRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Could not reach Anthropic API: {ex.Message}", ex);
            }

            var responseBody = await httpResponse.Content.ReadAsStringAsync();

            if (!httpResponse.IsSuccessStatusCode)
            {
                string claudeError;
                try
                {
                    using var errDoc = JsonDocument.Parse(responseBody);
                    claudeError = errDoc.RootElement
                        .GetProperty("error")
                        .GetProperty("message")
                        .GetString() ?? responseBody;
                }
                catch
                {
                    claudeError = $"HTTP {(int)httpResponse.StatusCode}";
                }

                throw new InvalidOperationException($"Claude API error: {claudeError}");
            }

            string extractedText;
            try
            {
                using var doc = JsonDocument.Parse(responseBody);
                var contentArray = doc.RootElement.GetProperty("content");
                extractedText = string.Empty;

                foreach (var block in contentArray.EnumerateArray())
                {
                    if (block.GetProperty("type").GetString() == "text")
                    {
                        extractedText = block.GetProperty("text").GetString() ?? string.Empty;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to parse Claude response: {ex.Message}", ex);
            }

            var clean = extractedText
                .Trim()
                .TrimStart('`')
                .TrimEnd('`');

            if (clean.StartsWith("json", StringComparison.OrdinalIgnoreCase))
                clean = clean[4..].Trim();

            try
            {
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var result = System.Text.Json.JsonSerializer.Deserialize<ExtractedFormDataDto>(clean, options);
                return result ?? new ExtractedFormDataDto();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to parse extraction result: {ex.Message}", ex);
            }
        }

        private byte[] GenerateDocxReport(ExtractedFormDataDto data)
        {
            string templateDir = Path.Combine("wwwroot", "templates");
            string templatePath = Path.Combine(templateDir, $"{data.ReportTemplate}.docx");

            if (!System.IO.File.Exists(templatePath) && System.IO.File.Exists(Path.Combine(templateDir, $"{data.ReportTemplate}.dotx")))
            {
                templatePath = Path.Combine(templateDir, $"{data.ReportTemplate}.dotx");
            }

            if (!System.IO.File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Template file not found at {templatePath}");
            }

            byte[] templateBytes = System.IO.File.ReadAllBytes(templatePath);

            using (var ms = new MemoryStream())
            {
                ms.Write(templateBytes, 0, templateBytes.Length);
                ms.Position = 0;

                using (var wordDoc = WordprocessingDocument.Open(ms, true))
                {
                    if (wordDoc.DocumentType == DocumentFormat.OpenXml.WordprocessingDocumentType.Template)
                    {
                        wordDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                    }

                    // 1. Map Simple Text Fields across Body, Headers, AND Footers

                    AppendTextToAllMatches(wordDoc, "DATE:", data.DOS);
                    AppendTextToAllMatches(wordDoc, "PATIENT:", $"{data.LastName}, {data.FirstName}");
                    AppendTextToAllMatches(wordDoc, "DATE OF BIRTH:", data.DOB);
                    AppendTextToAllMatches(wordDoc, "PATIENT MRN:", data.MRN);
                    string assistant = data.Assistant switch
                    {
                        "GD" => "George Davydov, P.A.",
                        "IA" => "Iraida Alvarez",
                        _ => data.Assistant
                    };
                    AppendTextToAllMatches(wordDoc, "ASSISTANT:", $"{assistant}");

                    string anesthesia = data.Anesthesia switch
                    {
                        "Gen." => "General",
                        "Sed." => "sedation",
                        "Block" => "Block",
                        _ => data.Anesthesia
                    };
                    AppendTextToAllMatches(wordDoc, "ANESTHESIA:", $"{anesthesia}");

                    OverwriteTextToAllMatches(wordDoc, "Ln, Fn", $"{data.LastName}, {data.FirstName}");
                    OverwriteTextToAllMatches(wordDoc, "(mmd/dd/yyyy)", $"{data.DOS}");

                    // Fetch all body paragraphs for the structural edits
                    var bodyParagraphs = wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();

                    // 2. Map Multi-line Bulleted Fields

                    var preOpHeading = bodyParagraphs.FirstOrDefault(p => p.InnerText.StartsWith("PREOPERATIVE DIAGNOSES:"));
                    //InsertNumberedList(preOpHeading, data.PreOp, useNumbering: true, startNumber: 1);
                    InsertNumberedListpre(preOpHeading, data.PreOp);

                    // 2. POSTOPERATIVE DIAGNOSES (Starts at 2, appending AFTER the default item)
                    var postOpHeading = bodyParagraphs.FirstOrDefault(p => p.InnerText.StartsWith("POSTOPERATIVE DIAGNOSES:"));
                    Paragraph postOpAnchor = postOpHeading;


                    string[] diagnoses = data.ICD10Rows.Split(',');
                    List<string> cleanDiagnoses = new List<string>();

                    foreach (string item in diagnoses)
                    {
                        string trimmedItem = item.Trim();
                        int firstSpaceIndex = trimmedItem.IndexOf(' ');

                        if (firstSpaceIndex > 0)
                        {
                            // Add only the description to the list
                            cleanDiagnoses.Add(trimmedItem.Substring(firstSpaceIndex + 1));
                        }
                        else
                        {
                            cleanDiagnoses.Add(trimmedItem);
                        }
                    }

                    // Combine the list back into a single comma-separated string
                    string finalString = string.Join(", ", cleanDiagnoses);
                    //InsertNumberedList(postOpAnchor, finalString, useNumbering: true, startNumber: 1);
                    InsertNumberedListpre(postOpAnchor, finalString);

                    // 3. PROCEDURES PERFORMED 
                    var procHeading = bodyParagraphs.FirstOrDefault(p => p.InnerText.StartsWith("PROCEDURES PERFORMED:"));
                    Paragraph procAnchor = procHeading;

                    string[] cpt = data.CPTRows.Split(',');
                    List<string> cleancpts = new List<string>();

                    foreach (string item in cpt)
                    {
                        string trimmedItem = item.Trim();
                        int firstSpaceIndex = trimmedItem.IndexOf(' ');

                        if (firstSpaceIndex > 0)
                        {
                            // Add only the description to the list
                            cleancpts.Add(trimmedItem.Substring(firstSpaceIndex + 1));
                        }
                        else
                        {
                            cleancpts.Add(trimmedItem);
                        }
                    }

                    // Combine the list back into a single comma-separated string
                    string finalStringcpt = string.Join(", ", cleancpts);

                    if (procHeading != null && procHeading.NextSibling<Paragraph>() != null)
                    {
                        procAnchor = procHeading.NextSibling<Paragraph>();
                    }
                    InsertNumberedList(procAnchor, finalStringcpt, useNumbering: true, startNumber: 2);
                    // InsertNumberedList(procAnchor, data.CPTRows, useNumbering: true, startNumber: 1);


                    // 3. Map ROM and Instability placeholders
                    //   ProcessRomAndInstability(bodyParagraphs, data.PreOpPassiveRoM);

                    if (!string.IsNullOrWhiteSpace(data.PreOpPassiveRoM))
                    {
                        var romDict = data.PreOpPassiveRoM.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(part => part.Split(new[] { ':' }, 2))
                            .Where(parts => parts.Length == 2)
                            .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim());

                        string ff = romDict.ContainsKey("FF") ? romDict["FF"] : "___";
                        string er = romDict.ContainsKey("ER") ? romDict["ER"] : "__";
                        string ir = romDict.ContainsKey("IR") ? romDict["IR"] :
                                   (romDict.ContainsKey("IRSS") ? romDict["IRSS"] : "__");
                        string instability = romDict.ContainsKey("Instability") ? romDict["Instability"] : "";

                        OverwriteTextToAllMatches(wordDoc, "#FF", $"{ff}");
                        OverwriteTextToAllMatches(wordDoc, "#ER ", $"{er}");
                        OverwriteTextToAllMatches(wordDoc, "#IR ", $"{ir}");
                        OverwriteTextToAllMatches(wordDoc, "#IN ", $"{instability}");

                    }



                    string bodypart = data.ReportTemplate switch
                    {
                        "RKAS" => "knee",
                        "RSAS" => "shoulder",
                        "LKAS" => "knee",
                        "LSAS" => "shoulder",
                        _ => data.ReportTemplate
                    };


                    // 4. Loop through them
                    foreach (var p in bodyParagraphs)
                    {
                        // Make sure the paragraph hasn't already been deleted
                        if (p.Parent == null) continue;

                        // --- HANDLE MVA ---
                        if (p.InnerText.Contains("[MVA]"))
                        {
                            // Grab the hardcoded sentence that sits on the very next line
                            var nextParagraph = p.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                            if (data.CaseType == "MVA")
                            {
                                // 1. Inject our newly generated text into the [MVA] tag
                                string mvaText = $"The patient’s {bodypart} injuries requiring surgery are the result of a motor vehicle accident.";
                                //p.InnerXml = p.InnerXml.Replace("[MVA]", mvaText);

                                //// 2. Delete the old hardcoded sentence below it so it doesn't duplicate!
                                //if (nextParagraph != null) nextParagraph.Remove();

                                // 1. Create the formatting: Times New Roman, Size 12 (24 half-points)
                                var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                                    new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" },
                                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                                );

                                // 2. Create the formatted text run
                                var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties, new DocumentFormat.OpenXml.Wordprocessing.Text(mvaText));

                                // 3. Clear the old tag and add the new formatted text
                                p.RemoveAllChildren();
                                p.AppendChild(formattedRun);

                                // 4. Delete the old hardcoded sentence below it so it doesn't duplicate!
                                if (nextParagraph != null) nextParagraph.Remove();
                            }
                            else
                            {
                                // If it is NOT an MVA case, delete the [MVA] tag AND the hardcoded sentence
                                p.Remove();
                                if (nextParagraph != null) nextParagraph.Remove();
                            }
                        }

                        // --- HANDLE WC ---
                        else if (p.InnerText.Contains("[WC]"))
                        {
                            // Grab the hardcoded sentence that sits on the very next line
                            var nextParagraph = p.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                            if (data.CaseType == "WC")
                            {
                                // 1. Inject our newly generated text into the [WC] tag
                                string wcText = $"The patient’s {bodypart} injuries requiring surgery are the result of an accident/incident that occurred at work.";
                                //p.InnerXml = p.InnerXml.Replace("[WC]", wcText);

                                //// 2. Delete the old hardcoded sentence below it so it doesn't duplicate!
                                //if (nextParagraph != null) nextParagraph.Remove();

                                // 1. Create the formatting: Times New Roman, Size 12 (24 half-points)
                                var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                                    new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" },
                                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                                );

                                // 2. Create the formatted text run
                                var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties, new DocumentFormat.OpenXml.Wordprocessing.Text(wcText));

                                // 3. Clear the old tag and add the new formatted text
                                p.RemoveAllChildren();
                                p.AppendChild(formattedRun);

                                // 4. Delete the old hardcoded sentence below it so it doesn't duplicate!
                                if (nextParagraph != null) nextParagraph.Remove();
                            }
                            else
                            {
                                // If it is NOT a WC case, delete the [WC] tag AND the hardcoded sentence
                                p.Remove();
                                if (nextParagraph != null) nextParagraph.Remove();
                            }
                        }
                    }


                    // modify the template based on the glenohumeral and subacromial. 
                    string[] GlenohumeralConditions = {
    "Adhesive capsulitis",
    "Biceps tendon",
    "Chondral lesion",
    "Partial RC tear",
    //"Partial RC tear",
    "Labrum tear",
    "SLAP tear, Type",
    "Loose bodies",
    "Synovitis of shoulder"
};

                    //foreach (var datagle in GlenohumeralConditions)
                    //{
                    //    if (data.ICD10Rows.Contains(datagle))
                    //    {
                    //        //keep row. 
                    //    }
                    //    else
                    //    {
                    //        //remove row
                    //    }
                    //}

                    string[] subacromialConditions = {
    "Bursitis",
    "complete rupture of rc",
    "Lysis/resection",
};

                    //foreach (var datagle in subacromialConditions)
                    //{
                    //    if (data.ICD10Rows.Contains(datagle) || data.ICD10Rows.Contains(datagle))
                    //    {
                    //        //keep row. 
                    //    }
                    //    else
                    //    {
                    //        //remove row
                    //    }
                    //}
                    //foreach (var p in bodyParagraphs)
                    //{
                    //    // Make sure the paragraph hasn't already been deleted
                    //    if (p.Parent == null) continue;

                    //    // --- HANDLE SYNOVECTOMY ---
                    //    if (p.InnerText.Contains("[Synovectomy]"))
                    //    {
                    //        // Grab the Labral Debridement Statement paragraph (next sibling)
                    //        var nextParagraph = p.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

                    //        // If Synovectomy is present, KEEP it and REMOVE Labral Debridement
                    //        if (p.InnerText.Contains("Synovectomy"))
                    //        {
                    //            // Format Synovectomy text
                    //            string synovectomyText = "Excessive glenohumeral synovitis was cleared with motorized shaver. The hypertrophic, erythematous synovium was resected, hemostasis was maintained with the radiofrequency device.";

                    //            var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    //                new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" },
                    //                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                    //            );

                    //            var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties,
                    //                new DocumentFormat.OpenXml.Wordprocessing.Text(synovectomyText));

                    //            // Replace the [Synovectomy] tag with formatted text
                    //            p.RemoveAllChildren();
                    //            p.AppendChild(formattedRun);

                    //            // Remove Labral Debridement Statement if it exists
                    //            if (nextParagraph != null && nextParagraph.InnerText.Contains("[Labral Debridement Statement]"))
                    //                nextParagraph.Remove();
                    //        }
                    //        else
                    //        {
                    //            // If Synovectomy is NOT present, remove both Synovectomy and Labral Debridement
                    //            p.Remove();
                    //            if (nextParagraph != null && nextParagraph.InnerText.Contains("[Labral Debridement Statement]"))
                    //                nextParagraph.Remove();
                    //        }
                    //    }
                    //}





                    // 4. Important: Explicitly save Header and Footer streams
                    foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                    {
                        headerPart.Header.Save();
                    }
                    foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                    {
                        footerPart.Footer.Save();
                    }

                    wordDoc.Save();
                }

                return ms.ToArray();
            }
        }
        private void OverwriteTextToAllMatches(WordprocessingDocument wordDoc, string searchText, string replaceText)
        {
            if (string.IsNullOrEmpty(searchText)) return;

            // Fallback to empty string if the incoming data is null
            string safeReplaceText = replaceText ?? string.Empty;

            var allParagraphs = new List<Paragraph>();

            // 1. Collect from Body
            if (wordDoc.MainDocumentPart.Document.Body != null)
            {
                allParagraphs.AddRange(wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>());
            }

            // 2. Collect from Headers
            foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
            {
                allParagraphs.AddRange(headerPart.Header.Descendants<Paragraph>());
            }

            // 3. Collect from Footers
            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                allParagraphs.AddRange(footerPart.Footer.Descendants<Paragraph>());
            }

            // 4. Find paragraphs that contain the search text anywhere inside them
            // This perfectly bypasses the split-text issue because InnerText reads the whole sentence
            var targetParas = allParagraphs.Where(p => p.InnerText.Contains(searchText)).ToList();

            foreach (var para in targetParas)
            {
                // Get the full glued-together text and apply the replacement
                string fullText = para.InnerText;
                string updatedText = fullText.Replace(searchText, safeReplaceText);

                // Try to preserve the paragraph's formatting (bold, font size, etc.)
                // by grabbing the properties of the first run before we delete the fragments
                var firstRun = para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();
                var runProperties = firstRun?.RunProperties?.CloneNode(true);

                // Clear the old, fragmented text runs out of the paragraph
                para.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Run>();

                // Rebuild the paragraph with the clean, updated text wrapped in a single Run
                var newRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
                if (runProperties != null)
                {
                    newRun.AppendChild(runProperties);
                }

                var textNode = new DocumentFormat.OpenXml.Wordprocessing.Text(updatedText)
                {
                    Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
                };

                // 2. Define the formatting: Times New Roman, Size 12 (24 half-points)
                var runProperties1 = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.RunFonts()
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        ComplexScript = "Times New Roman"
                    },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                );

                // 3. Create the Run, passing in the properties first, then the text
                var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties1, textNode);

                // 4. Append the formatted text to the paragraph
                newRun.AppendChild(formattedRun);

                //newRun.AppendChild(textNode);
                para.AppendChild(newRun);
            }
        }
        private void AppendTextToAllMatches(WordprocessingDocument wordDoc, string prefix, string appendText)
        {
            if (string.IsNullOrWhiteSpace(appendText)) return;

            var allParagraphs = new List<Paragraph>();

            // Collect from Body
            if (wordDoc.MainDocumentPart.Document.Body != null)
            {
                allParagraphs.AddRange(wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>());
            }

            // Collect from Headers
            foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
            {
                allParagraphs.AddRange(headerPart.Header.Descendants<Paragraph>());
            }

            // Collect from Footers
            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                allParagraphs.AddRange(footerPart.Footer.Descendants<Paragraph>());
            }

            // Find ALL matches so if "PATIENT:" is in the body AND header, both get filled out
            var targetParas = allParagraphs.Where(p => p.InnerText.Trim().StartsWith(prefix)).ToList();

            foreach (var para in targetParas)
            {
                // 1. Create the text node
                var textNode = new DocumentFormat.OpenXml.Wordprocessing.Text($" {appendText}")
                {
                    Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
                };
                //para.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(textNode));

                // 2. Define the formatting: Times New Roman, Size 12 (24 half-points)
                var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.RunFonts()
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        ComplexScript = "Times New Roman"
                    },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                );

                // 3. Create the Run, passing in the properties first, then the text
                var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties, textNode);

                // 4. Append the formatted text to the paragraph
                para.AppendChild(formattedRun);
            }
        }

        private void InsertNumberedList(Paragraph anchorPara, string csvData, bool useNumbering = true, int startNumber = 1)
        {
            if (string.IsNullOrWhiteSpace(csvData) || anchorPara == null) return;

            var items = csvData.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                               .Select(i => i.Trim())
                               .ToList();

            Paragraph currentPara = anchorPara;
            for (int i = 0; i < items.Count; i++)
            {
                var newPara = new Paragraph();

                // Apply numbering only if requested, starting from the designated number
                string textContent = useNumbering ? $"{startNumber + i}. {items[i]}" : items[i];

                var textNode = new DocumentFormat.OpenXml.Wordprocessing.Text(textContent)
                {
                    Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
                };

                //newPara.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(textNode));

                // 2. Define the formatting: Times New Roman, Size 12 (24 half-points)
                var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.RunFonts()
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        ComplexScript = "Times New Roman"
                    },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                );

                // 3. Create the Run, passing in the properties first, then the text
                var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties, textNode);

                // 4. Append the formatted text to the paragraph
                newPara.AppendChild(formattedRun);


                // Insert directly into the parent container to support Lists inside Tables
                currentPara.Parent.InsertAfter(newPara, currentPara);

                // Move the anchor forward so the next item goes below this one
                currentPara = newPara;
            }
        }

        private void InsertNumberedListpre(Paragraph anchorPara, string csvData)
        {
            if (string.IsNullOrWhiteSpace(csvData) || anchorPara == null) return;

            var items = csvData.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                               .Select(i => i.Trim())
                               .ToList();

            Paragraph currentPara = anchorPara;
            for (int i = 0; i < items.Count; i++)
            {
                var newPara = new Paragraph();
                var textNode = new DocumentFormat.OpenXml.Wordprocessing.Text($"{i + 1}. {items[i]}")
                {
                    Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
                };

                // newPara.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(textNode));

                // 2. Define the formatting: Times New Roman, Size 12 (24 half-points)
                var runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.RunFonts()
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        ComplexScript = "Times New Roman"
                    },
                    new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" }
                );

                // 3. Create the Run, passing in the properties first, then the text
                var formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run(runProperties, textNode);

                // 4. Append the formatted text to the paragraph
                newPara.AppendChild(formattedRun);


                // Insert directly into the parent container to support Lists inside Tables
                currentPara.Parent.InsertAfter(newPara, currentPara);

                currentPara = newPara;
            }
        }

        private void ProcessRomAndInstability(List<Paragraph> bodyParagraphs, string preOpPassiveRom)
        {
            if (string.IsNullOrWhiteSpace(preOpPassiveRom)) return;

            var romDict = preOpPassiveRom.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Split(new[] { ':' }, 2))
                .Where(parts => parts.Length == 2)
                .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim());

            string ff = romDict.ContainsKey("FF") ? romDict["FF"] : "___";
            string er = romDict.ContainsKey("ER") ? romDict["ER"] : "__";
            string ir = romDict.ContainsKey("IR") ? romDict["IR"] :
                       (romDict.ContainsKey("IRSS") ? romDict["IRSS"] : "__");
            string instability = romDict.ContainsKey("Instability") ? romDict["Instability"] : "";

            var romPara = bodyParagraphs.FirstOrDefault(p => p.InnerText.Contains("passive range of motion"));
            if (romPara != null)
            {
                romPara.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Run>();
                string newText = $"A preoperative orthopedic examination revealed a passive range of motion of {ff} degrees of forward elevation, {er} degrees of external rotation, and {ir} degrees of internal rotation.";
                romPara.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text(newText) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
            }

            var instPara = bodyParagraphs.FirstOrDefault(p => p.InnerText.Contains("Stability examination revealed:"));
            if (instPara != null)
            {
                instPara.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Run>();
                string newText = $"Stability examination revealed: {instability} instability.";
                instPara.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text(newText) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
            }
        }

        private static string SanitizeFilename(string input)
        {
            var invalid = Path.GetInvalidFileNameChars();
            return string.Concat(input.Select(c => invalid.Contains(c) ? '_' : c));
        }

        #endregion

    }
}
public class PdfExtractRequest
{
    public List<string> Pages { get; set; } = new(); // base64 PNG strings, one per page
    public string FileName { get; set; } = string.Empty;
}

public class ExtractedFormDataDto
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string DOB { get; set; }
    public string DOS { get; set; }
    public string MRN { get; set; }
    public string Assistant { get; set; }
    public string Anesthesia { get; set; }
    public string LocationLH { get; set; }
    public string CaseType { get; set; }
    public string ReportTemplate { get; set; }
    public string CPTRows { get; set; }
    public string ICD10Rows { get; set; }
    public string IntraOpRows { get; set; }
    public string PreOp { get; set; }
    public string PreOpPassiveRoM { get; set; }
}