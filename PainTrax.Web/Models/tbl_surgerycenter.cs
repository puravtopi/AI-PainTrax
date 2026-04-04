using System.ComponentModel.DataAnnotations;
namespace PainTrax.Web.Models
{
	public class tbl_surgerycenter
    {
       
        [Required(ErrorMessage = "Please enter Name")]
        public string? Surgerycenter_name { get; set; }
       
        public string? Address { get; set; }
		public string? ContactNo { get; set; }
        public string? ContactPerson { get; set; }
        public int? Id { get; set; }
        public int? cmp_id { get; set; }
        public bool? is_active { get; set; } = true;

    }
}