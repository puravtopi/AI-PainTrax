namespace PainTrax.Web.Models
{
    public class PatientDetails
    {
        public string id { get; set; }
        public string sign { get; set; }
        public string fname { get; set; }
        public string lname { get; set; }
        public DateTime dob { get; set; }
        public bool isExist { get; set; }
    }
}
