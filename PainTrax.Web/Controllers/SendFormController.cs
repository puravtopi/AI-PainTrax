using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    public class SendFormController : Controller
    {
        private readonly CommonService _commonservices = new CommonService();
        private readonly PdfHelper _pdfhelper = new PdfHelper();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<HomeController> _logger;

        public SendFormController(Microsoft.AspNetCore.Hosting.IHostingEnvironment environment, ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            Environment = environment;
            _logger = logger;
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult AuthoAckno(string id)
        {
            id = Uri.UnescapeDataString(id);
            id = id.Replace(" ", "+");
            var _id = EncryptionHelper.Decrypt(id);

            //if (_commonservices.isSignExist(_id) == false)
            //{
            var data = _commonservices.getPatientDetails(_id);

            if (data == null)
            {
                data = new Models.PatientDetails()
                {
                    fname = "",
                    lname = ""

                };
            }
            data.id = _id;
            data.isExist = false;
            return View(data);
            //}
            //else
            //{
            //    var data = new Models.PatientDetails()
            //    {
            //        isExist = true
            //    };
            //    return View(data);
            //}

        }

        public IActionResult AuthoAcknoKDVPC(string id)
        {
            id = Uri.UnescapeDataString(id);
            id = id.Replace(" ", "+");
            var _id = EncryptionHelper.Decrypt(id);

            //if (_commonservices.isSignExist(_id) == false)
            //{
            var data = _commonservices.getPatientDetails(_id);

            if (data == null)
            {
                data = new Models.PatientDetails()
                {
                    fname = "",
                    lname = ""

                };
            }
            data.id = _id;
            data.isExist = false;
            return View(data);
            //}
            //else
            //{
            //    var data = new Models.PatientDetails()
            //    {
            //        isExist = true
            //    };
            //    return View(data);
            //}

        }

        [HttpPost]
        public IActionResult SaveSign([FromBody] PatientDetails model)
        {
            if (string.IsNullOrEmpty(model.signatureData))
                return BadRequest("Invalid signature data.");
            try
            {

                var base64Data = ExtractBase64Data(model.signatureData);

                if (base64Data == null || !IsBase64String(base64Data))
                {
                    return BadRequest("Invalid Base-64 string.");
                }


                var imageData = Convert.FromBase64String(base64Data);
                var signaturesDir = Path.Combine(Environment.WebRootPath, "signatures");


                if (!Directory.Exists(signaturesDir))
                {
                    Directory.CreateDirectory(signaturesDir);
                }
                var filename = $"{model.id}.jpeg";
                var savePath = Path.Combine(signaturesDir, filename);

                System.IO.File.WriteAllBytes(savePath, imageData);

                var id = _commonservices.InsertSign(filename, base64Data, model.id.ToString());

                var client_code = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
                var pdffilename = "PATIENT MASTER AUTHORIZATION.pdf";
                if (client_code.ToLower() == "kdvpc")
                    pdffilename = "CF_KDV.pdf";

                var source = Path.Combine(Environment.WebRootPath, "Autorization", pdffilename);
                Dictionary<string, string> controls = new Dictionary<string, string>();

                //Get Patient Data with Sign
                var pData = _commonservices.getPatientDetailsWithSign(model.id.ToString());

                if (pData != null)
                {
                    controls.Add("dob", pData.dob.Value.ToString("MM-dd-yyyy"));
                    controls.Add("patientName", pData.lname + ' ' + pData.fname);
                    controls.Add("executionDate", model.executionDate?.ToString("MM-dd-yyyy"));
                    controls.Add("primaryPhone", model.primaryPhone);
                    controls.Add("alternatePhone", model.alternatePhone);
                    controls.Add("email", model.email);
                    controls.Add("patientPortalUsername", model.patientPortalUsername);
                    controls.Add("SMSConsentY", string.IsNullOrEmpty(model.sMSConsent) ? ""
                    : char.ToUpper(model.sMSConsent.ToLowerInvariant()[0]) + model.sMSConsent.ToLowerInvariant().Substring(1));
                    controls.Add("SMSConsentN", string.IsNullOrEmpty(model.sMSConsent) ? ""
                    : char.ToUpper(model.sMSConsent.ToLowerInvariant()[0]) + model.sMSConsent.ToLowerInvariant().Substring(1));

                    controls.Add("EmailConsentY", string.IsNullOrEmpty(model.emailConsent) ? ""
                    : char.ToUpper(model.emailConsent.ToLowerInvariant()[0]) + model.emailConsent.ToLowerInvariant().Substring(1));
                    controls.Add("EmailConsentN", string.IsNullOrEmpty(model.emailConsent) ? ""
                    : char.ToUpper(model.emailConsent.ToLowerInvariant()[0]) + model.emailConsent.ToLowerInvariant().Substring(1));

                    controls.Add("VoiceMailConsentY", string.IsNullOrEmpty(model.voiceMailConsent) ? ""
                    : char.ToUpper(model.voiceMailConsent.ToLowerInvariant()[0]) + model.voiceMailConsent.ToLowerInvariant().Substring(1));
                    controls.Add("VoiceMailConsentN", string.IsNullOrEmpty(model.voiceMailConsent) ? ""
                    : char.ToUpper(model.voiceMailConsent.ToLowerInvariant()[0]) + model.voiceMailConsent.ToLowerInvariant().Substring(1));

                    controls.Add("PatientPortalConsentY", string.IsNullOrEmpty(model.patientPortalConsent) ? ""
                    : char.ToUpper(model.patientPortalConsent.ToLowerInvariant()[0]) + model.patientPortalConsent.ToLowerInvariant().Substring(1));
                    controls.Add("PatientPortalConsentN", string.IsNullOrEmpty(model.patientPortalConsent) ? ""
                    : char.ToUpper(model.patientPortalConsent.ToLowerInvariant()[0]) + model.patientPortalConsent.ToLowerInvariant().Substring(1));

                    controls.Add("AudioRecordingConsentY", string.IsNullOrEmpty(model.audioRecordingConsent) ? ""
                    : char.ToUpper(model.audioRecordingConsent.ToLowerInvariant()[0]) + model.audioRecordingConsent.ToLowerInvariant().Substring(1));
                    controls.Add("AudioRecordingConsentN", string.IsNullOrEmpty(model.audioRecordingConsent) ? ""
                    : char.ToUpper(model.audioRecordingConsent.ToLowerInvariant()[0]) + model.audioRecordingConsent.ToLowerInvariant().Substring(1));

                    controls.Add("MarketingConsentY", string.IsNullOrEmpty(model.marketingConsent) ? ""
                    : char.ToUpper(model.marketingConsent.ToLowerInvariant()[0]) + model.marketingConsent.ToLowerInvariant().Substring(1));
                    controls.Add("MarketingConsentN", string.IsNullOrEmpty(model.marketingConsent) ? ""
                    : char.ToUpper(model.marketingConsent.ToLowerInvariant()[0]) + model.marketingConsent.ToLowerInvariant().Substring(1));

                    controls.Add("SignedByName", model.signedByName);
                    controls.Add("SignedDate", model.signedDate?.ToString("MM-dd-yyyy"));
                    controls.Add("RelationshipToPatient", model.relationshipToPatient);

                    string path = Path.Combine(_env.ContentRootPath, "PatientDocuments", "Others", model.id.ToString());

                    _pdfhelper.AcknoStamping(source, controls, path, "PATIENT MASTER AUTHORIZATION.pdf", savePath);
                }

                return Ok(new { id = id });
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SaveSign");
                return Ok(new { id = 0 });
            }
        }

        private string ExtractBase64Data(string signatureData)
        {
            // Example method to extract Base64 data from a comma-separated string
            // Adjust as per your actual data structure
            return signatureData.Contains(",")
                   ? signatureData.Split(',')[1].Trim()
                   : signatureData.Trim();
        }

        private bool IsBase64String(string base64)
        {
            Span<byte> buffer = new Span<byte>(new byte[base64.Length]);
            return Convert.TryFromBase64String(base64, buffer, out _);
        }

        #region private Method
        private void SaveLog(Exception ex, string actionname)
        {
            var msg = "";
            if (ex.InnerException == null)
            {
                _logger.LogError(ex.Message);
                msg = ex.Message;
            }
            else
            {
                _logger.LogError(ex.InnerException.Message);
                msg = ex.InnerException.Message;
            }
            var logdata = new tbl_log
            {
                CreatedDate = DateTime.Now,
                CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId),
                Message = msg,
            };
            new LogService().Insert(logdata);
        }
        #endregion 
    }
}
