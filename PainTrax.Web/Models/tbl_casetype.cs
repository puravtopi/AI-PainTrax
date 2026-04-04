using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
    public class tbl_casetype
    {
        [Required(ErrorMessage = "Please enter CaseType")]
        public string? casetype { get; set; }
        public int Id { get; set; }       
        public int? cmp_id { get; set; }
    }
}
