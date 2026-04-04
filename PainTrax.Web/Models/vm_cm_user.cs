namespace PainTrax.Web.Models
{
    public class vm_cm_user
    {
        public string? fname { get; set; }
        public string? lname { get; set; }
        public string emailid { get; set; }
        public string? address { get; set; }
        public string? group_name { get; set; }

        public string? desig_name { get; set; }
        public int? Id { get; set; }
     
        public string? fullname { get; set; }  
      
     
        public string uname { get; set; }
        public string password { get; set; }
        public int? groupid { get; set; }
        public int? desigid { get; set; }
        public int? cmp_id { get; set; }
        public string company_name { get; set; }
        public string client_code { get; set; }
        public bool? is_newuser { get; set; }

    }
}
