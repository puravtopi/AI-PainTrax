using System.ComponentModel.DataAnnotations;
namespace PainTrax.Web.Models
{
    public class tbl_inscos
    {
        public string? contactpersonname { get; set; }

        [Required(ErrorMessage = "Please enter Insco Name ")]
        public string? cmpname { get; set; }
        public string? emailid { get; set; }
        public string? telephone { get; set; }
        public string? address1 { get; set; }
        public int? id { get; set; }
        public string? address2 { get; set; }
        public string? city { get; set; }
        public string? state { get; set; }
        public string? zipcode { get; set; }
        public string? faxno { get; set; }
        public bool? setasdefault { get; set; }
        public bool isactive { get; set; }
        public int? cmp_id { get; set; }
        public DateTime? createddate { get; set; }
        public int? createdby { get; set; }
        public DateTime? updatedate { get; set; }
        public int? updatedby { get; set; }
        public int? old_id { get; set; }
        public string? label { get; set; }
        public string? val { get; set; }
        public int? ival { get; set; }

    }
}