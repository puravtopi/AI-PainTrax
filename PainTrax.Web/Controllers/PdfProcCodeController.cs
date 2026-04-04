namespace PainTrax.Web.Controllers
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using Microsoft.AspNetCore.Mvc;
    using Org.BouncyCastle.Asn1.Ocsp;
    using PainTrax.Web.Helper;
    using PainTrax.Web.Models;
    using PainTrax.Web.Services;
    using System.Data;

    public class PdfProcCodeController : Controller
    {
        private readonly PdfProcCodeService _service;

        public PdfProcCodeController()
        {
            _service = new PdfProcCodeService();
        }

        // ----------------------------------------------------
        // LIST
        // ----------------------------------------------------
        public IActionResult Index()
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            var list = _service.GetAll(" and cmp_id ="+cmpid);
            return View(list);
        }

        // ----------------------------------------------------
        // CREATE GET
        // ----------------------------------------------------
        public IActionResult Create()
        {
            return View(new tbl_pdfproccode());
        }

        // ----------------------------------------------------
        // CREATE POST
        // ----------------------------------------------------
        [HttpPost]
        public IActionResult Create(tbl_pdfproccode model)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            if (ModelState.IsValid)
            {
                model.icdcodes = string.Join(",", new[] { model.icdcode1, model.icdcode2, model.icdcode3, model.icdcode4 }.Where(x => !string.IsNullOrWhiteSpace(x)));
                model.cptcodes = string.Join(",", new[] { model.cptcode1, model.cptcode2, model.cptcode3, model.cptcode4 }.Where(x => !string.IsNullOrWhiteSpace(x)));
                model.cmp_id = Convert.ToInt32( cmpid);
                _service.Insert(model);
                return RedirectToAction("Index");
            }
            return View(model);
        }

        // ----------------------------------------------------
        // EDIT GET
        // ----------------------------------------------------
        public IActionResult Edit(int id)
        {
            var data = _service.GetOne(id);
            if (data == null) return NotFound();
            return View(data);
        }

        // ----------------------------------------------------
        // EDIT POST
        // ----------------------------------------------------
        [HttpPost]
        public IActionResult Edit(tbl_pdfproccode model)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            if (ModelState.IsValid)
            {
                model.icdcodes = string.Join(",", new[] { model.icdcode1, model.icdcode2, model.icdcode3, model.icdcode4 }.Where(x => !string.IsNullOrWhiteSpace(x)));
                model.cptcodes = string.Join(",", new[] { model.cptcode1, model.cptcode2, model.cptcode3, model.cptcode4 }.Where(x => !string.IsNullOrWhiteSpace(x)));
                model.cmp_id = Convert.ToInt32(cmpid);

                _service.Update(model);
                return RedirectToAction("Index");
            }
            return View(model);
        }

        // ----------------------------------------------------
        // DELETE GET
        // ----------------------------------------------------
        public IActionResult Delete(int id)
        {
            var data = _service.GetOne(id);
            if (data == null) return NotFound();
            return View(data);
        }

        // ----------------------------------------------------
        // DELETE POST
        // ----------------------------------------------------
        [HttpPost]
        public IActionResult ConfirmDelete(int id)
        {
            _service.Delete(id);
            return RedirectToAction("Index");
        }

        // ----------------------------------------------------
        // CSV IMPORT GET
        // ----------------------------------------------------
        public IActionResult Import()
        {
            return View();
        }

        
        private string GetCellValue(Cell cell, SharedStringTablePart stringTable)
        {
            if (cell == null || cell.CellValue == null)
                return "";

            string value = cell.CellValue.InnerText;

            // Shared string?
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            // Assume 'value' is a string variable holding "MAC_x000D_"

            // Assign the result of the replacement back to the 'value' variable
           // value = value.Replace("_x000D_", "\n");

            // The 'value' variable now contains "MAC\n"
            return value;


        }

    // ----------------------------------------------------
    // Excel IMPORT POST
    // ----------------------------------------------------

    [HttpPost]
        public IActionResult ImportXlsx(IFormFile file)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            if (file == null || file.Length == 0)
                return RedirectToAction("Index");

            using var stream = file.OpenReadStream();
            using var doc = SpreadsheetDocument.Open(stream, false);

            var workbookPart = doc.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            var stringTable = workbookPart.SharedStringTablePart;

            bool skipHeader = true;

            foreach (Row row in sheetData.Elements<Row>())
            {
                if (skipHeader) { skipHeader = false; continue; }

                var cells = row.Elements<Cell>().ToList();

                string Get(int i) =>
                    i < cells.Count ? GetCellValue(cells[i], stringTable) : "";

                DataTable dt = _service.GetData("select * from tbl_pdfproccode where mcode='" + Get(1) + "' and cmp_id=" + cmpid);

                if (dt.Rows.Count == 0)
                {
                    var model = new tbl_pdfproccode
                    {
                        mcode = Get(1),
                        mprocedure = Get(2),
                        cptcodes = Get(3),
                        icdcodes = Get(4),
                        specialequ = Get(5),
                        diagnosis = Get(6),
                        speequchk = Get(7),
                        mprocshort = Get(8),
                        cptcode1 = Get(9),
                        cptcode2 = Get(10),
                        cptcode3 = Get(11),
                        cptcode4 = Get(12),
                        icdcode1 = Get(13),
                        icdcode2 = Get(14),
                        icdcode3 = Get(15),
                        icdcode4 = Get(16),
                        cmp_id = Convert.ToInt32(cmpid),
                        cmp_code = ""
                    };

                    // Rebuild icdcodes from valid values only
                    model.icdcodes = string.Join(",",
                        new[] { model.icdcode1, model.icdcode2, model.icdcode3, model.icdcode4 }
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                    );
                    model.cptcodes = string.Join(",", new[] { model.cptcode1, model.cptcode2, model.cptcode3, model.cptcode4 }.Where(x => !string.IsNullOrWhiteSpace(x)));

                    _service.Insert(model);
                }
            }

            
            return RedirectToAction("Index");
        }



        // ----------------------------------------------------
        // CSV IMPORT POST
        // ----------------------------------------------------
       [HttpPost]
        public IActionResult ImportCsv(IFormFile file)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            if (file == null || file.Length == 0)
            {
                ViewBag.Error = "Please upload CSV";
                return View("Import");
            }

            using var reader = new StreamReader(file.OpenReadStream());

            string? line;

            // Skip header
            reader.ReadLine();

            while ((line = reader.ReadLine()) != null)
            {
             
                var parts = ParseCsvLine(line);

                // Ensure at least 17 fields
                while (parts.Count < 17) parts.Add("");

                var model = new tbl_pdfproccode
                {
                    mcode = parts[1].Trim(),
                    mprocedure = parts[2].Trim(),
                    cptcodes = parts[3].Trim(),
                    icdcodes = parts[4].Trim(),
                    specialequ = parts[5].Trim(),
                    diagnosis = parts[6].Trim(),
                    speequchk = parts[7].Trim(),
                    mprocshort = parts[8].Trim(),
                    cptcode1 = parts[9].Trim(),
                    cptcode2 = parts[10].Trim(),
                    cptcode3 = parts[11].Trim(),
                    cptcode4 = parts[12].Trim(),
                    icdcode1 = parts[13].Trim(),
                    icdcode2 = parts[14].Trim(),
                    icdcode3 = parts[15].Trim(),
                    icdcode4 = parts[16].Trim(),

                    cmp_id = 18,
                    cmp_code = ""
                };

                _service.Insert(model);
            }

        
            return RedirectToAction("Index");
        }

        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            string value = "";

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '\"')
                {
                    // Toggle quote status
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    // Comma ends value
                    result.Add(value);
                    value = "";
                }
                else
                {
                    value += c;
                }
            }

            result.Add(value); // add last field
            return result;
        }

    }

}
