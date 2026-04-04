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


            var sdata = _surgeryCentreService.GetAll(cnd);
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

            var surgoryList = _surgeryCentreService.GetAll(cnd);

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
                    dt.Rows.Add(cnt.doe?.ToShortDateString() ?? "", cnt.location,cnt.WC,cnt.NF,cnt.LIEN , cnt.NoOFIE, cnt.NoOFFU);
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
                    dt.Rows.Add(t.name, t.sex, t.mc, t.casetype, t.location, t.vaccinated, t.mcode, t.bodypart, t.ins_ver_status, t.mc_status, t.case_status, t.insverstatus, t.vac_status, t.scheduled, t.executed, t.requested, t.DOE == null ? "" : t.DOE.Value.ToShortDateString(),t.PolicyNo,t.Attorney);//, IVFR.scheduled == null ? "" : IVFR.scheduled.Value.ToShortDateString());
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



    }
}
