namespace PainTrax.Web.Models
{
    public class tbl_dictation
    {
        public int? id { get; set; }

        public string? voice_file { get; set; }
        public string? txt_file { get; set; }
        public string? type { get; set; }
        public int? ie_id { get; set; }
        public int? fu_id { get; set; }
        public int? cmp_id { get; set; }
    }
}
