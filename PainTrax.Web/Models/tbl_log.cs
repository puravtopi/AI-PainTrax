namespace PainTrax.Web.Models
{
    public class tbl_log
    {
        public int? Id { get; set; }
        public DateTime? CreatedDate { get; set; }
        public int? CreatedBy { get; set; }
        public string Message { get; set; }
    }
}
