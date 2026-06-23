namespace PainTrax.Web.Models
{
    public class tbl_intake_forward_setting
    {
        public int Id { get; set; }
        public bool History { get; set; }
        public bool Cc { get; set; }
        public bool Pe { get; set; }
        public bool Diagnosis { get; set; }
        public bool Neroexam { get; set; }
        public bool Adl { get; set; }
        public bool Note { get; set; }
        public int Cmp_id { get; set; }
    }
}
