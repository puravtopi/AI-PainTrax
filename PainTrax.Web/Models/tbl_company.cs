using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{
    public class tbl_company
    {
        [Required(ErrorMessage = "Please Enter  Name")]
        public string? name { get; set; }
        public string? address { get; set; }
        public string? email { get; set; }
        public string? phone { get; set; }
        public bool is_active { get; set; }
        public int? id { get; set; }
        public string? username { get; set; }
        public string? password { get; set; }
        public string? client_code { get; set; }     
        public int? created_by { get; set; }
        public DateTime? created_date { get; set; }
        public string? cmp_type { get; set; }

    }
}