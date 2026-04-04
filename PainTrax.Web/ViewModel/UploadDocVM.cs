namespace PainTrax.Web.ViewModel
{
    public class UploadDocVM
    {
        public string DirName { get; set; }
        public int? CmpId { get; set; }

        public List<IFormFile> files { get; set; }
    }
}
