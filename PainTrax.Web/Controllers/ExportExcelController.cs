using Microsoft.AspNetCore.Mvc;
using MS.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PainTrax.Services;
using PainTrax.Web.Helper;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    public class ExportExcelController : Controller
    {
        private readonly ILogger<FormsController> _logger;

        private readonly IWebHostEnvironment _environment;
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly PatientService _patientservices = new PatientService();
        private readonly ParentService _pareentservices = new ParentService();
        private readonly Common _commonservices = new Common();
        private readonly InscosService _inscosservices = new InscosService();
        private readonly AttorneysService _attorneyservices = new AttorneysService();
        private readonly AadjusterService _aadjusterService = new AadjusterService();
        private readonly EmpService _empService = new EmpService();
        private readonly UserService _userService = new UserService();

        public ExportExcelController(ILogger<FormsController> logger, IWebHostEnvironment environment)
        {
            _environment = environment;
            _logger = logger;
        }


        public IActionResult Index(string id)
        {
            using var stream = new MemoryStream();
            string fname = "", lname = "";
            var patient = _patientservices.GetOne(Convert.ToInt32(id));

            var iepatient = _ieService.GetOnebyPatientIdNew(patient?.id ?? 0);
            fname = patient?.fname ?? "";
            lname = patient?.lname ?? "";
            var priins = _inscosservices.GetOne(iepatient?.primary_ins_cmp_id ?? 0);
            var secins = _inscosservices.GetOne(iepatient?.secondary_ins_cmp_id ?? 0);
            var att = _attorneyservices.GetOne(iepatient?.attorney_id ?? 0);
            var emp = _empService.GetOne(iepatient?.emp_id ?? 0);
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Create(stream,
                DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                SheetData sheetData = new SheetData();

                // Create Header Row
                Row headerRow = new Row();
                string[] columns =
          {
                "Case Type",
                "First Name",
                "Middle Name",
                "Last Name",
                "Suffix",
                "DateOfBirth",
                "Patient Address Line 1",
                "Patient Address Line 2",
                "Patient City",
                "Patient State",
                "Patient Zip",
                "Accident Date",
                "Accident State",
                "Gender",
                "SSN",
                "Marrital Status",
                "Home Phone",
                "Work Phone",
                "Work Phone Ext",
                "Cell Phone",
                "Patient Email",
                "Insurance",
                "Insurance Address 1",
                "Insurance Address 2",
                "Insurance City",
                "Insurance State",
                "Insurance Zip",
                "Insurance Phone",
                "Insurance Fax",
                "Insurance Contact Person",
                "Policy#",
                "Claim#",
                "Wcb#",
                "Member Id",
                "Group Id",
                "Policy Type",
                "Sec Insurance",
                "Sec Insurance Address 1",
                "Sec Insurance Address 2",
                "Sec Insurance City",
                "Sec Insurance State",
                "Sec Insurance Zip",
                "Sec Insurance Phone",
                "Sec Insurance Fax",
                "Sec Insurance Contact Person",
                "Sec Policy#",
                "Sec Claim#",
                "Policy Holder",
                "Policy Holder Suffix",
                "Policy Holder SSN#",
                "Policy Holder Address 1",
                "Policy Holder Address 2",
                "Policy Holder City",
                "Policy Holder State",
                "Policy Holder Zip",
                "Policy Holder Home Phone",
                "Policy Holder Work Phone",
                "Policy Holder Work Phone Ext",
                "Policy Holder Cell",
                "Policy Holder Relation to Patient",
                "Policy Holder Gender",
                "Policy Holder DOB",
                "Emergency First Name",
                "Emergency Middle Name",
                "Emergency Last Name",
                "Emergency Suffix",
                "Emergency Address 1",
                "Emergency Address 2",
                "Emergency City",
                "Emergency State",
                "Emergency Zip",
                "Emergency Home Phone",
                "Emergency Cell Phone",
                "Emergency Work Phone",
                "Emergency WorkExt",
                "Emergency Email",
                "Emergency Relation To Patient",
                "Date Of First Treatement",
                "Employer Name",
                "Empoyer Address",
                "Employer City",
                "Employer state",
                "Employer Zip",
                "Empoyer Phone",
                "Attorney Name",
                "Reffering Office Name",
                "Reffering Doctor Name",
                "Reffering Doctor Npi",
                "ExternalPatientId"
            };

                foreach (var col in columns)
                {
                    headerRow.Append(CreateCell(col));
                }

                sheetData.Append(headerRow);
                Row valueRow = new Row();

                string[] values =
                 {
                  iepatient?.compensation ?? "", // "Case Type",

                    patient?.fname ?? "", // "First Name",
                    patient?.mname ?? "", // "Middle Name",
                    patient?.lname ?? "", // "Last Name",

                    patient?.gender == "1" ? "Mr" :
                    patient?.gender == "2" ? "Ms" : "", // "Suffix",

                    patient?.dob?.ToString("dd/MM/yyyy") ?? "", // "DateOfBirth",

                    patient?.address ?? "", // "Patient Address Line 1",
                    "", // "Patient Address Line 2",
                    patient?.city ?? "", // "Patient City",
                    patient?.state ?? "", // "Patient State",
                    patient?.zip ?? "", // "Patient Zip",

                    iepatient?.doa?.ToString("dd/MM/yyyy") ?? "", // "Accident Date",
                    iepatient?.state ?? "", // "Accident State",

                    patient?.gender == "1" ? "Male" :
                    patient?.gender == "2" ? "Female" : "Other", // "Gender",

                    patient?.ssn ?? "", // "SSN",
                    "", // "Marrital Status",

                    patient?.home_ph ?? "", // "Home Phone",
                    "", // "Work Phone",
                    "", // "Work Phone Ext",

                    patient?.mobile ?? "", // "Cell Phone",
                    patient?.email ?? "", // "Patient Email",

                    // Primary Insurance
                    priins?.cmpname ?? "", // "Insurance",
                    priins?.address1 ?? "", // "Insurance Address 1",
                    "", // "Insurance Address 2",
                    priins?.city ?? "", // "Insurance City",
                    priins?.state ?? "", // "Insurance State",
                     "", // "Insurance Zip",
                    priins?.telephone ?? "", // "Insurance Phone",
                    priins?.faxno ?? "", // "Insurance Fax",
                    "", // "Insurance Contact Person",

                    iepatient?.primary_policy_no ?? "", // "Policy#",
                    iepatient?.primary_claim_no ?? "", // "Claim#",

                    iepatient?.primary_wcb_group ?? "", // "Wcb#",
                    "", // "Member Id",

                     "", // "Group Id",
                    "", // "Policy Type",

                    // Secondary Insurance
                    secins?.cmpname ?? "", // "Sec Insurance",
                    secins?.address1 ?? "", // "Sec Insurance Address 1",
                    "", // "Sec Insurance Address 2",
                    secins?.city ?? "", // "Sec Insurance City",
                    secins?.state ?? "", // "Sec Insurance State",
                    "", // "Sec Insurance Zip",
                    secins?.telephone ?? "", // "Sec Insurance Phone",
                    secins?.faxno ?? "", // "Sec Insurance Fax",
                    "", // "Sec Insurance Contact Person",

                    iepatient?.secondary_policy_no ?? "", // "Sec Policy#",
                    iepatient?.secondary_claim_no ?? "", // "Sec Claim#",

                    // Policy Holder
                    (patient?.fname ?? "") + " " + (patient?.lname ?? ""), // "Policy Holder",

                    patient?.gender == "1" ? "Mr" :
                    patient?.gender == "2" ? "Ms" : "", // "Policy Holder Suffix",

                    patient?.ssn ?? "", // "Policy Holder SSN#",
                    patient?.address ?? "", // "Policy Holder Address 1",
                    "", // "Policy Holder Address 2",
                    patient?.city ?? "", // "Policy Holder City",
                    patient?.state ?? "", // "Policy Holder State",
                    patient?.zip ?? "", // "Policy Holder Zip",

                    patient?.home_ph ?? "", // "Policy Holder Home Phone",
                    "", // "Policy Holder Work Phone",
                    "", // "Policy Holder Work Phone Ext",

                    patient?.mobile ?? "", // "Policy Holder Cell",
                    "Self", // "Policy Holder Relation to Patient",

                    patient?.gender == "1" ? "Male" :
                    patient?.gender == "2" ? "Female" : "Other", // "Policy Holder Gender",

                    patient?.dob?.ToString("dd/MM/yyyy") ?? "", // "Policy Holder DOB",

                    // Emergency Contact
                    "", // "Emergency First Name",
                    "", // "Emergency Middle Name",
                    "", // "Emergency Last Name",
                    "", // "Emergency Suffix",
                    "", // "Emergency Address 1",
                    "", // "Emergency Address 2",
                    "", // "Emergency City",
                    "", // "Emergency State",
                    "", // "Emergency Zip",
                    "", // "Emergency Home Phone",
                    "", // "Emergency Cell Phone",
                    "", // "Emergency Work Phone",
                    "", // "Emergency WorkExt",
                    "", // "Emergency Email",
                    "", // "Emergency Relation To Patient",

                    // Treatment
                    iepatient?.doe?.ToString("dd/MM/yyyy") ?? "", // "Date Of First Treatement",

                    // Employer
                    emp?.name ?? "", // "Employer Name",
                    emp?.address ?? "", // "Empoyer Address",
                    "", // "Employer City",
                    "", // "Employer state",
                    "", // "Employer Zip",

                    emp?.phone ?? "", // "Empoyer Phone",

                    // Attorney
                    att?.Attorney ?? "", // "Attorney Name",

                    // Referring
                    "", // "Reffering Office Name",

                    "", // "Reffering Doctor Name",
                    "", // "Reffering Doctor Npi",

                    // External Patient Id
                    patient?.id?.ToString() ?? "" // "ExternalPatientId"
                };

       

                foreach (var val in values)
                {
                    valueRow.Append(CreateCell(val));
                }

                sheetData.Append(valueRow);

                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Patients"
                };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();
            }

            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{lname}_{fname}.xlsx"
            );
        }

        private Cell CreateCell(string text)
        {
            return new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };
        }
    }
}
