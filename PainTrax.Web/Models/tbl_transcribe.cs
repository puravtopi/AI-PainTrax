namespace PainTrax.Web.Models
{
    public class tbl_transcribe
    {
        public int id { get; set; }
        public string content { get; set; }
        public string history { get; set; }
        public string cc { get; set; }
        public string plan { get; set; }
        public string subjective { get; set; }
        public string assessment { get; set; }
        public long ie_id { get; set; }
        public long fu_id { get; set; }
        public int length { get; set; }
        public int intake_id { get; set; }
        public int cmp_id { get; set; }
        public DateTime date { get; set; }
    }
}
