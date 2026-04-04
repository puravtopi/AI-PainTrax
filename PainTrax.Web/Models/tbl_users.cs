using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models

{
    public class tbl_users
    {

        public int? Id { get; set; }
        [Required(ErrorMessage = "Please Enter First Name")]
        public string? fname { get; set; }
        [Required(ErrorMessage = "Please Enter Last Name")]
        public string? lname { get; set; }

        [Required(ErrorMessage = "Please Enter Email")]
        public string? emailid { get; set; }
        public string? address { get; set; }
        public string? fullname { get; set; }
        public string? phoneno { get; set; }

        [Required(ErrorMessage = "Please Enter UserName")]
        public string? uname { get; set; }

        //[Required(ErrorMessage = "Please Enter Password")]
        [Required]
        [DataType(DataType.Password)]
        [RegularExpression(@"^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*\W).{8,}$",
        ErrorMessage = "Password must be at least 8 characters and include uppercase, lowercase, number, and special character.")]
        public string? password { get; set; }

        [Required(ErrorMessage = "Please Select Group")]
        public int? groupid { get; set; }

        [Required(ErrorMessage = "Please Select Designation")]
        public int? desigid { get; set; }
        public int? cmp_id { get; set; }
        public int? createdby { get; set; }
        public DateTime? createddate { get; set; }
        public DateTime? updateddate { get; set; }
        public int? updatedby { get; set; }
        public string? signature { get; set; }
        public string? signature_hidden { get; set; }
        public string? providername { get; set; }
        public string? assistant_providername { get; set; }
        public bool? is_newuser { get; set; }

    }
}