using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
    public class tbl_gender
    {
        public int Id { get; set; }
        [Required(ErrorMessage = "Please Select Gender!")]
        public string Gender { get; set; }
        public int? cmp_id { get; set; }
    }
}
