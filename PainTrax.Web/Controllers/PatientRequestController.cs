using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;

namespace PainTrax.Web.Controllers
{
    public class PatientRequestController : Controller
    {
        #region variables
        private readonly Common _commonservices = new Common();
        private readonly IntakeService _intakeservice = new IntakeService();
        private readonly IEmailService _emailService;
        #endregion

        #region Ctor
        public PatientRequestController(IEmailService emailService)
        {
            _emailService = emailService;
        }
        #endregion

        public IActionResult Index()
        {
            var cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);
            var data = _intakeservice.GetInitialAllIntakeAI(cmpid.Value);
            return View(data);
        }

        [HttpPost]
        public async Task<IActionResult> SendRequest(PatientRequestVM model)
        {

            if (model != null)
            {
                var cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                InitialIntakeAI initialIntakeAI = new InitialIntakeAI()
                {
                    Id = 0,
                    CmpId = cmpid,
                    Visit_Type = "IE",
                    FN = model.FirstName,
                    LN = model.LastName,
                    Email = model.Email,
                    Mobile = model.Mobile,
                    SendDate = System.DateTime.Now,
                    LocationId = model.LocationId,
                };
                var result = _intakeservice.SaveInitialIntakeAI(initialIntakeAI);
                int Id;
                bool isValid = int.TryParse(result, out Id);
                if (isValid)
                {
                    var email = model.Email;
                    if (email != "null" && email != "")
                    {
                        //var locid = EncryptionHelper.Encrypt(model.LocationId.ToString());
                        //var id = EncryptionHelper.Encrypt(Id.ToString());
                        var locid = model.LocationId.ToString();
                        var id = Id.ToString();
                        string link = "https://www.paintrax.com/v2/IntakeForm/PatientInitialIntake?locid=" + locid + "&id=" + id;
                        var subject = "Intake Form Required Before Your Visit";

                        var body = System.IO.File.ReadAllText("wwwroot/Uploads/EmailTemplate/PatientRequestForm.html")
                                       .Replace("{RESET_LINK}", link);

                        await _emailService.SendEmailAsync(email, subject, body);

                        return Json("1");
                    }
                }
            }
            return Json("1");
        }
    }
}
