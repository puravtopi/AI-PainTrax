namespace PainTrax.Web.Models
{
    public class tbl_referring_physician
    {
        public int Id { get; set; }
        public string? physicianname { get; set; }
        public string? address{ get; set; }
        public int? locationid {  get; set; }
        public int? cmp_id {  get; set; }
    }
}
