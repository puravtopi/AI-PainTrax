namespace PainTrax.Web.ViewModel
{
    public class PatientRequestVM
    {
        public int LocationId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Mobile { get; set; }
        public DateTime SendDate { get; set; }
        public DateTime? SubmitDate { get; set; }
    }
}
