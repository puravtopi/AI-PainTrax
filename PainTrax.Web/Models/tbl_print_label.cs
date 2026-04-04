namespace PainTrax.Web.Models
{
    public class tbl_print_label
    {
        public int id { get; set; }
        public int cmp_id { get; set; }
        public string lbl_code { get; set; }
        public string lbl_title { get; set; }
        public bool? is_show { get; set; }
        public bool? is_new_line { get; set; }
        public string? style { get; set; }
    }
}
