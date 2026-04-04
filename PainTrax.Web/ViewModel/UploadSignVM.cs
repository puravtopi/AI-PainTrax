namespace PainTrax.Web.ViewModel
{
    public class UploadSignVM
    {
        public IList<IFormFile> signs { get; set; }
        public IFormFile patient { get; set; }
        public IFormFile patientFU { get; set; }
    }
}
