namespace PainTrax.Web.Models
{
    public class tbl_ie_sign
    {
        public int id { get; set; }
        public string? type { get; set; }
        public int? patient_id { get; set; }
        public string? signatureData { get; set; }
        public string? signatureValue { get; set; }
    }
}
