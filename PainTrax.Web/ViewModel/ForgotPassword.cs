using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.ViewModel
{
    public class ForgotPassword
    {
        [Required(ErrorMessage = "Please Enter Email")]
        public string email {  get; set; }

        [Required(ErrorMessage = "Please Enter Client Code")]
        public string companycode { get; set; }
    }
}
