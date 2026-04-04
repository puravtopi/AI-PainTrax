namespace PainTrax.Web.Models
{
    public class tbl_pdfproccode
    {
        public int id { get; set; }
        public string mcode { get; set; } = string.Empty;
        public string? mprocedure { get; set; }
        public string? cptcodes { get; set; }
        public string? icdcodes { get; set; }
        public string? specialequ { get; set; }
        public string? diagnosis { get; set; }
        public string? speequchk { get; set; }
        public string? mprocshort { get; set; }
        public string? cptcode1 { get; set; }
        public string? cptcode2 { get; set; }
        public string? cptcode3 { get; set; }
        public string? cptcode4 { get; set; }
        public string? icdcode1 { get; set; }
        public string? icdcode2 { get; set; }
        public string? icdcode3 { get; set; }
        public string? icdcode4 { get; set; }
        public int? cmp_id { get; set; }
        public string? cmp_code { get; set; }
    }
}
