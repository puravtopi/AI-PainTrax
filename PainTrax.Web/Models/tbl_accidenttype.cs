using System.ComponentModel.DataAnnotations;
namespace PainTrax.Web.Models
{
    public class tbl_accidenttype
    {       
        [Required(ErrorMessage = "Please enter AccidentType")]
        public string? accidenttype { get; set; }
        public int Id { get; set; }
        public int? cmp_id { get; set; }
    }
}
