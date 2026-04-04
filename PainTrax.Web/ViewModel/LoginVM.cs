namespace PainTrax.Web.ViewModel
{
    public class LoginVM
    {
        public int Id { get; set; }
        public string UserName { get; set; }
        public string PassWord { get; set; }
        public string ClientCode { get; set; }
        public bool? is_newuser { get; set; } = false;
    }
}
