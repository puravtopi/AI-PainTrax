using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.ViewModel
{
    public class ChangePassword
    {
        public string password { get; set; }
        public string Token { get; set; }
        public string cmpid { get; set; }
        
        [Required]
        [DataType(DataType.Password)]
        [RegularExpression(@"^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*\W).{8,}$",
        ErrorMessage = "Password must be at least 8 characters and include uppercase, lowercase, number, and special character.")]
        public string? Password { get; set; }


        [DataType(DataType.Password)]
        [Display(Name = "Confirm password")]
        [RegularExpression(@"^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*\W).{8,}$",
        ErrorMessage = "Password must be at least 8 characters and include uppercase, lowercase, number, and special character.")]

        [Compare("Password", ErrorMessage = "The password and confirmation password do not match.")]
        public string? ConfirmPassword { get; set; }
    }
}
