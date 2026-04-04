using System.ComponentModel.DataAnnotations;
namespace PainTrax.Web.Models
{
    public class tbl_state
    {
        [Required(ErrorMessage = "Please Enter State Name")]
        public string? state_name { get; set; }
        public string? fullname { get; set; }
        public int id { get; set; }
        public int cmp_id { get; set; }
    //    public string? label { get; set; }
    //    public string? val { get; set; }        
    }
}
