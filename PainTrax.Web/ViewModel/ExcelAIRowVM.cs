using Newtonsoft.Json;          // or System.Text.Json — swap as needed
using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.ViewModel
{
    // PdfExtractRequest is already defined in your project — no duplicate needed.

    /// <summary>One surgical case row — matches the editable preview table columns.</summary>
    public class ExcelAIRowVM
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public string DOB { get; set; } = "";
        public string DOS { get; set; } = "";
        public string MRN { get; set; } = "";
        public string Assistant { get; set; } = "";
        public string Anesthesia { get; set; } = "";
        public string LocationLH { get; set; } = "";
        public string CaseType { get; set; } = "";
        public string ReportTemplate { get; set; } = "";
        public string CPTRows { get; set; } = "";
        public string ICD10Rows { get; set; } = "";
        public string IntraOpRows { get; set; } = "";
    }
}
