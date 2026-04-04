using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
    public class tbl_signature
    {
        public int id { get; set; }
        [Required(ErrorMessage= "Please enter FirstName")]
        public string fname { get; set; }
        [Required(ErrorMessage = "Please enter LastName")]
        public string lname { get; set; }
        [Required(ErrorMessage = "Please enter Dob")]
        public string dob { get; set; }
        public string? signaturePath { get; set; }
        public int? cmp_id { get; set; }
    }
}
