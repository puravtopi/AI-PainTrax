using System.ComponentModel.DataAnnotations;

namespace PainTrax.Web.Models
{

    public class tbl_provider
    {

        public int? id { get; set; }

        [Required(ErrorMessage ="Please Enter Provider")]
        public string? provider { get; set; }

        [Required(ErrorMessage = "Please Enter Provider EmailId")]
        public string? emailaddress { get; set; }
        public string? telephone { get; set; }
        public string? contactperson { get; set; }
        public string? address1 { get; set; }
        public string? address2 { get; set; }
        public string? city { get; set; }
        public string? state { get; set; }
        public string? zip { get; set; }       
        public bool? set_as_default { get; set; }
        public string? fax { get; set; }
        public int? cmp_id { get; set; }
        public int? created_by { get; set; }
        public int? updated_by { get; set; }
        public DateTime? created_date { get; set; }
        public DateTime? updated_date { get; set; }

    }
}