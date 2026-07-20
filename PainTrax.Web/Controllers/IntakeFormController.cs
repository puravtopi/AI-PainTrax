using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GroupDocs.Viewer.Results;
using HtmlToOpenXml;
using MailKit;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using MS.Models;
using MS.Services;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Asn1.Sec;
using Org.BouncyCastle.Utilities.Collections;
using PainTrax.Services;
using PainTrax.Web.AzureServices;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using System.Diagnostics.Metrics;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static PainTrax.Web.Helper.EnumHelper;


namespace PainTrax.Web.Controllers
{
    //[SessionCheckFilter]
    public class IntakeFormController : Controller
    {
        #region Vriables
        private readonly Common _commonservices = new Common();
        private readonly IntakeService service = new IntakeService();
        private readonly LocationsService _locservices = new LocationsService();
        private readonly ProviderService _povservices = new ProviderService();
        private readonly POCServices _pocservices = new POCServices();
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly PatientService _patientservices = new PatientService();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        private readonly IWebHostEnvironment _env;
        private readonly AzureAIServices _azureService;
        private readonly DiagcodesService _diagcodesService = new DiagcodesService();
        private readonly TreatmentMasterService _treatmentService = new TreatmentMasterService();
        private readonly ILogger<IntakeFormController> _logger;
        private readonly UserService _userService = new UserService();
        private readonly DefaultValueSettingServices _defaultSettingService = new DefaultValueSettingServices();
        private readonly SettingsService _settingservices = new SettingsService();
        private readonly DefaultDataServices _defaultService = new DefaultDataServices();
        #endregion

        public IntakeFormController(
         Microsoft.AspNetCore.Hosting.IHostingEnvironment environment,
         IWebHostEnvironment env, AzureAIServices azureService,
         ILogger<IntakeFormController> logger
        )
        {
            Environment = environment;
            _env = env;
            _azureService = azureService;
            _logger = logger;

        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Create()
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            IntakeForm obj = new IntakeForm();
            return View(obj);
        }
        [HttpPost]
        public IActionResult Create(IntakeForm model)
        {

            string what_test = "", medical_condition = "", social_history = "", symptoms_since_accident = "",
                cc_neck = "", cc_neck_radiates = "", cc_neck_tingling = "", cc_neck_increase = "", cc_midback = "", cc_midback_increase = "",
                cc_lowback = "", cc_lowback_radiates = "", cc_lowback_tingling = "", cc_lowback_increase = "", cc_l_shoulder = "", cc_l_shoulder_increase = "",
                cc_r_shoulder = "", cc_r_shoulder_increase = "", cc_l_knee = "", cc_l_knee_increase = "",
                cc_r_knee = "", cc_r_knee_increase = "", cc_other_1 = "", cc_other_2 = "";

            if (model.what_tests_xray == "true")
                what_test = what_test + "," + model.what_tests_xray;
            if (model.what_tests_ct == "true")
                what_test = what_test + "," + model.what_tests_ct;
            if (model.what_tests_mri == "true")
                what_test = what_test + "," + model.what_tests_mri;

            model.what_test = what_test.TrimStart(',');

            if (model.any_medical_conditions_Diabeties == "true")
                medical_condition = medical_condition + "," + model.any_medical_conditions_Diabeties;
            if (model.any_medical_conditions_bp == "true")
                medical_condition = medical_condition + "," + model.any_medical_conditions_bp;
            if (model.any_medical_conditions_ashthma == "true")
                medical_condition = medical_condition + "," + model.any_medical_conditions_ashthma;
            if (model.any_medical_conditions_heart == "true")
                medical_condition = medical_condition + "," + model.any_medical_conditions_heart;
            if (model.any_medical_conditions_none == "true")
                medical_condition = medical_condition + "," + model.any_medical_conditions_none;

            model.any_medical_conditions = medical_condition.TrimStart(',');

            if (!string.IsNullOrEmpty(model.smoke))
            {
                if (model.smoke == "Yes")
                {
                    if (!string.IsNullOrEmpty(model.txt_smoke))
                        social_history = social_history + ", Smoke for " + model.txt_smoke;
                    else
                        social_history = social_history + ", Smoke";
                }
                else if (model.smoke == "No")
                {
                    social_history = social_history + ", No Smoke";
                }
            }

            if (!string.IsNullOrEmpty(model.marijuana))
            {
                if (model.marijuana == "Yes")
                {
                    if (!string.IsNullOrEmpty(model.txt_marijuana))
                        social_history = social_history + ", marijuana for " + model.txt_marijuana;
                    else
                        social_history = social_history + ", marijuana";
                }
                else if (model.marijuana == "No")
                {
                    social_history = social_history + ", No marijuana";
                }
            }

            if (!string.IsNullOrEmpty(model.marijuana))
            {
                if (model.marijuana == "Yes")
                {
                    if (!string.IsNullOrEmpty(model.txt_marijuana))
                        social_history = social_history + ", marijuana for " + model.txt_marijuana;
                    else
                        social_history = social_history + ", marijuana";
                }
                else if (model.marijuana == "No")
                {
                    social_history = social_history + ", No marijuana";
                }
            }

            if (!string.IsNullOrEmpty(model.alcohol))
            {
                if (model.alcohol == "Yes")
                {
                    if (!string.IsNullOrEmpty(model.alcohol))
                        social_history = social_history + ", alcohol for " + model.alcohol;
                    else
                        social_history = social_history + ", alcohol";
                }
                else if (model.alcohol == "No")
                {
                    social_history = social_history + ", No alcohol";
                }
            }


            model.social_history = social_history.TrimStart(',');


            if (model.symptoms_of_accident_Headaches == "true")
                symptoms_since_accident = symptoms_since_accident + ",Headaches";
            if (model.symptoms_of_accident_ChestPain == "true")
                symptoms_since_accident = symptoms_since_accident + ",Chest Pain/Short of Breath";
            if (model.symptoms_of_accident_Abdominal == "true")
                symptoms_since_accident = symptoms_since_accident + ",Abdominal Pain";
            if (model.symptoms_of_accident_Muscle == "true")
                symptoms_since_accident = symptoms_since_accident + ",Muscle Spasms";
            if (model.symptoms_of_accident_Dizziness == "true")
                symptoms_since_accident = symptoms_since_accident + ",Dizziness";
            if (model.symptoms_of_accident_Nausea == "true")
                symptoms_since_accident = symptoms_since_accident + ",Nausea";
            if (model.symptoms_of_accident_Ringing_in_ears == "true")
                symptoms_since_accident = symptoms_since_accident + ",Ringing in Ears";
            if (model.symptoms_of_accident_Bladder == "true")
                symptoms_since_accident = symptoms_since_accident + ",Bladder Incontinence";
            if (model.symptoms_of_accident_Bowel == "true")
                symptoms_since_accident = symptoms_since_accident + ",Bowel Incontinence";
            if (model.symptoms_of_accident_Seizure == "true")
                symptoms_since_accident = symptoms_since_accident + ",Seizure";
            if (model.symptoms_of_accident_Sleep_issues == "true")
                symptoms_since_accident = symptoms_since_accident + ",Sleep Issues/Difficulty";
            if (model.symptoms_of_accident_Anxiety == "true")
                symptoms_since_accident = symptoms_since_accident + ",Anxiety/Depression";


            model.symptoms_since_accident = symptoms_since_accident.TrimStart(',');

            if (!string.IsNullOrEmpty(model.txt_describe_neck))
                cc_neck = "The patient complains of neck pain that is " + model.txt_describe_neck + "/10, with 10 being the worst , which is ";

            if (model.describe_neck_Constant == "true")
                cc_neck = cc_neck + ",Constant";
            if (model.describe_neck_Intermittent == "true")
                cc_neck = cc_neck + ",Intermittent";
            if (model.describe_neck_Sharp == "true")
                cc_neck = cc_neck + ",Sharp";
            if (model.describe_neck_Electric == "true")
                cc_neck = cc_neck + ",Electric";
            if (model.describe_neck_Shooting == "true")
                cc_neck = cc_neck + ",Shooting";
            if (model.describe_neck_Throbbing == "true")
                cc_neck = cc_neck + ",Throbbing";
            if (model.describe_neck_Pulsating == "true")
                cc_neck = cc_neck + ",Pulsating";
            if (model.describe_neck_Dull == "true")
                cc_neck = cc_neck + ",Dull";
            if (model.describe_neck_Achy == "true")
                cc_neck = cc_neck + ",Achy";

            model.cc_neck = cc_neck.TrimStart(',') + ".";

            if (model.neck_pain_radiates_RUE == "true")
                cc_neck_radiates = cc_neck_radiates + ",RUE";
            if (model.neck_pain_radiates_LUE == "true")
                cc_neck_radiates = cc_neck_radiates + ",LUE";
            if (model.neck_pain_radiates_BUE == "true")
                cc_neck_radiates = cc_neck_radiates + ",BUE";
            if (model.neck_pain_numbness == "true")
                cc_neck_radiates = cc_neck_radiates + ",numbness";

            if (!string.IsNullOrEmpty(cc_neck_radiates))
                model.cc_neck = model.cc_neck + " Radiates To " + cc_neck_radiates.TrimStart(',') + ".";

            if (model.neck_pain_bodypart_shoulder == "true")
                cc_neck_tingling = cc_neck_tingling + ",shoulder";
            if (model.neck_pain_bodypart_elbow == "true")
                cc_neck_tingling = cc_neck_tingling + ",elbow";
            if (model.neck_pain_bodypart_hand == "true")
                cc_neck_tingling = cc_neck_tingling + ",hand";
            if (model.neck_pain_bodypart_wrist == "true")
                cc_neck_tingling = cc_neck_tingling + ",wrist";
            if (model.neck_pain_bodypart_finger == "true")
                cc_neck_tingling = cc_neck_tingling + ",finger";

            if (!string.IsNullOrEmpty(cc_neck_tingling))
                model.cc_neck = model.cc_neck + " Tingling To " + cc_neck_tingling.TrimStart(',') + ".";


            if (model.increase_neck_pain_lookingup == "true")
                cc_neck_increase = cc_neck_increase + ",looking up";
            if (model.increase_neck_pain_lookingdown == "true")
                cc_neck_increase = cc_neck_increase + ",looking down";
            if (model.increase_neck_pain_turningheadright == "true")
                cc_neck_increase = cc_neck_increase + ",turning head to right";
            if (model.increase_neck_pain_turningheadleft == "true")
                cc_neck_increase = cc_neck_increase + ",turning head to left";
            if (model.increase_neck_pain_driving == "true")
                cc_neck_increase = cc_neck_increase + ",driving";
            if (model.increase_neck_pain_twisting == "true")
                cc_neck_increase = cc_neck_increase + ",twisting";

            if (!string.IsNullOrEmpty(cc_neck_increase))
                model.cc_neck = model.cc_neck + " Pain increases by " + cc_neck_increase.TrimStart(',') + ".";

            if (!string.IsNullOrEmpty(model.txt_describe_midback))
                cc_midback = "The patient complains of midback pain that is " + model.txt_describe_midback + "/10, with 10 being the worst , which is ";


            if (model.describe_midback_Constant == "true")
                cc_midback = cc_midback + ",Constant";
            if (model.describe_midback_Intermittent == "true")
                cc_midback = cc_midback + ",Intermittent";
            if (model.describe_midback_Sharp == "true")
                cc_midback = cc_midback + ",Sharp";
            if (model.describe_midback_Electric == "true")
                cc_midback = cc_midback + ",Electric";
            if (model.describe_midback_Shooting == "true")
                cc_midback = cc_midback + ",Shooting";
            if (model.describe_midback_Throbbing == "true")
                cc_midback = cc_midback + ",Throbbing";
            if (model.describe_midback_Pulsating == "true")
                cc_midback = cc_midback + ",Pulsating";
            if (model.describe_midback_Dull == "true")
                cc_midback = cc_midback + ",Dull";
            if (model.describe_midback_Achy == "true")
                cc_midback = cc_midback + ",Achy";


            model.cc_midback = cc_midback.TrimStart(',');


            if (model.increase_midback_pain_sitting == "true")
                cc_midback_increase = cc_midback_increase + ",sitting";
            if (model.increase_midback_pain_standing == "true")
                cc_midback_increase = cc_midback_increase + ",standing";
            if (model.increase_midback_pain_bendingforward == "true")
                cc_midback_increase = cc_midback_increase + ",bending forward";
            if (model.increase_midback_pain_bendingbackwards == "true")
                cc_midback_increase = cc_midback_increase + ", bending backwards";
            if (model.increase_midback_pain_sleeping == "true")
                cc_midback_increase = cc_midback_increase + ",sleeping";
            if (model.increase_midback_pain_twisting == "true")
                cc_midback_increase = cc_midback_increase + ",twisting";
            if (model.increase_midback_pain_lifting == "true")
                cc_midback_increase = cc_midback_increase + ",lifting";





            if (!string.IsNullOrEmpty(cc_lowback_increase))
                model.cc_midback = model.cc_midback + " Pain increases by " + cc_midback_increase.TrimStart(',') + ".";


            if (!string.IsNullOrEmpty(model.txt_describe_lowback))
                cc_lowback = "The patient complains of lowback pain that is " + model.txt_describe_lowback + "/10, with 10 being the worst , which is ";

            if (model.describe_lowback_Constant == "true")
                cc_lowback = cc_lowback + ",Constant";
            if (model.describe_lowback_Intermittent == "true")
                cc_lowback = cc_lowback + ",Intermittent";
            if (model.describe_lowback_Sharp == "true")
                cc_lowback = cc_lowback + ",Sharp";
            if (model.describe_lowback_Electric == "true")
                cc_lowback = cc_lowback + ",Electric";
            if (model.describe_lowback_Shooting == "true")
                cc_lowback = cc_lowback + ",Shooting";
            if (model.describe_lowback_Throbbing == "true")
                cc_lowback = cc_lowback + ",Throbbing";
            if (model.describe_lowback_Pulsating == "true")
                cc_lowback = cc_lowback + ",Pulsating";
            if (model.describe_lowback_Dull == "true")
                cc_lowback = cc_lowback + ",Dull";
            if (model.describe_lowback_Achy == "true")
                cc_lowback = cc_lowback + ",Achy";


            model.cc_lowback = cc_lowback.TrimStart(',');




            if (model.lowback_pain_radiates_RLE == "true")
                cc_lowback_radiates = cc_lowback_radiates + ",RLE";
            if (model.lowback_pain_radiates_LLE == "true")
                cc_lowback_radiates = cc_lowback_radiates + ",LLE";
            if (model.lowback_pain_radiates_BLE == "true")
                cc_lowback_radiates = cc_lowback_radiates + ",BLE";
            if (model.lowback_pain_numbness == "true")
                cc_lowback_radiates = cc_lowback_radiates + ",numbness";

            if (!string.IsNullOrEmpty(cc_lowback_radiates))
                model.cc_lowback = model.cc_lowback + " Radiates To " + cc_lowback_radiates.TrimStart(',') + ".";

            if (model.lowback_pain_bodypart_thigh == "true")
                cc_lowback_tingling = cc_lowback_tingling + ",thigh";
            if (model.lowback_pain_bodypart_knee == "true")
                cc_lowback_tingling = cc_lowback_tingling + ",knee";
            if (model.lowback_pain_bodypart_leg == "true")
                cc_lowback_tingling = cc_lowback_tingling + ",leg";
            if (model.lowback_pain_bodypart_ankle == "true")
                cc_lowback_tingling = cc_lowback_tingling + ",ankle";
            if (model.lowback_pain_bodypart_foot == "true")
                cc_lowback_tingling = cc_lowback_tingling + ",foot";
            if (model.lowback_pain_bodypart_toe == "true")
                cc_lowback_tingling = cc_lowback_tingling + ",toe";


            if (!string.IsNullOrEmpty(cc_lowback_tingling))
                model.cc_lowback = model.cc_lowback + " Tingling To " + cc_lowback_tingling.TrimStart(',') + ".";


            if (model.increase_lowback_pain_sitting == "true")
                cc_lowback_increase = cc_lowback_increase + ",sitting";
            if (model.increase_lowback_pain_standing == "true")
                cc_lowback_increase = cc_lowback_increase + ",standing";
            if (model.increase_lowback_pain_bending_forward == "true")
                cc_lowback_increase = cc_lowback_increase + ",bending forward";
            if (model.increase_lowback_pain_bending_backwards == "true")
                cc_lowback_increase = cc_lowback_increase + ",bending backwards";
            if (model.increase_lowback_pain_sleeping == "true")
                cc_lowback_increase = cc_lowback_increase + ",sleeping";
            if (model.increase_lowback_pain_twisting_right == "true")
                cc_lowback_increase = cc_lowback_increase + ",twisting right";
            if (model.increase_lowback_pain_twisting_left == "true")
                cc_lowback_increase = cc_lowback_increase + ",twisting left";
            if (model.increase_lowback_pain_lifting == "true")
                cc_lowback_increase = cc_lowback_increase + ",lifting";



            if (!string.IsNullOrEmpty(cc_lowback_increase))
                model.cc_lowback = model.cc_lowback + " Pain increases by " + cc_lowback_increase.TrimStart(',') + ".";

            if (!string.IsNullOrEmpty(model.describe_leftshoulder))
                cc_l_shoulder = "The patient complains of left shoulder pain that is " + model.describe_leftshoulder + "/10, with 10 being the worst , which is ";


            if (model.txt_describe_leftshoulder_Constant == "true")
                cc_l_shoulder = cc_l_shoulder + ",Constant";
            if (model.txt_describe_leftshoulder_Intermittent == "true")
                cc_l_shoulder = cc_l_shoulder + ",Intermittent";
            if (model.txt_describe_leftshoulder_Sharp == "true")
                cc_l_shoulder = cc_l_shoulder + ",Sharp";
            if (model.txt_describe_leftshoulder_Electric == "true")
                cc_l_shoulder = cc_l_shoulder + ",Electric";
            if (model.txt_describe_leftshoulder_Shooting == "true")
                cc_l_shoulder = cc_l_shoulder + ",Shooting";
            if (model.txt_describe_leftshoulder_Throbbing == "true")
                cc_l_shoulder = cc_l_shoulder + ",Throbbing";
            if (model.txt_describe_leftshoulder_Pulsating == "true")
                cc_l_shoulder = cc_l_shoulder + ",Pulsating";
            if (model.txt_describe_leftshoulder_Dull == "true")
                cc_l_shoulder = cc_l_shoulder + ",Dull";
            if (model.txt_describe_leftshoulder_Achy == "true")
                cc_l_shoulder = cc_l_shoulder + ",Achy";


            model.cc_l_shoulder = cc_l_shoulder.TrimStart(',');



            if (model.increase_leftshoulder_pain_Raising_arm == "true")
                cc_l_shoulder_increase = cc_l_shoulder_increase + ",Raising arm";
            if (model.increase_leftshoulder_pain_Lifting == "true")
                cc_l_shoulder_increase = cc_l_shoulder_increase + ",Lifting";
            if (model.increase_leftshoulder_pain_Working == "true")
                cc_l_shoulder_increase = cc_l_shoulder_increase + ",Working";
            if (model.increase_leftshoulder_pain_Rotation == "true")
                cc_l_shoulder_increase = cc_l_shoulder_increase + ",Rotation";
            if (model.increase_leftshoulder_pain_Overhead_activities == "true")
                cc_l_shoulder_increase = cc_l_shoulder_increase + ",Overhead activities";


            if (!string.IsNullOrEmpty(cc_l_shoulder_increase))
                model.cc_l_shoulder = model.cc_l_shoulder + " Pain increases by " + cc_l_shoulder_increase.TrimStart(',') + ".";

            if (!string.IsNullOrEmpty(model.describe_rightshoulder))
                cc_r_shoulder = "The patient complains of right shoulder pain that is " + model.describe_rightshoulder + "/10, with 10 being the worst , which is ";


            if (model.txt_describe_rightshoulder_Constant == "true")
                cc_r_shoulder = cc_r_shoulder + ",Constant";
            if (model.txt_describe_rightshoulder_Intermittent == "true")
                cc_r_shoulder = cc_r_shoulder + ",Intermittent";
            if (model.txt_describe_rightshoulder_Sharp == "true")
                cc_r_shoulder = cc_r_shoulder + ",Sharp";
            if (model.txt_describe_rightshoulder_Electric == "true")
                cc_r_shoulder = cc_r_shoulder + ",Electric";
            if (model.txt_describe_rightshoulder_Shooting == "true")
                cc_r_shoulder = cc_r_shoulder + ",Shooting";
            if (model.txt_describe_rightshoulder_Throbbing == "true")
                cc_r_shoulder = cc_r_shoulder + ",Throbbing";
            if (model.txt_describe_rightshoulder_Pulsating == "true")
                cc_r_shoulder = cc_r_shoulder + ",Pulsating";
            if (model.txt_describe_rightshoulder_Dull == "true")
                cc_r_shoulder = cc_r_shoulder + ",Dull";
            if (model.txt_describe_rightshoulder_Achy == "true")
                cc_r_shoulder = cc_r_shoulder + ",Achy";


            model.cc_r_shoulder = cc_r_shoulder.TrimStart(',');

            if (model.increase_rightshoulder_pain_Raising_arm == "true")
                cc_r_shoulder_increase = cc_r_shoulder_increase + ",Raising arm";
            if (model.increase_rightshoulder_pain_Lifting == "true")
                cc_r_shoulder_increase = cc_r_shoulder_increase + ",Lifting";
            if (model.increase_rightshoulder_pain_Working == "true")
                cc_r_shoulder_increase = cc_r_shoulder_increase + ",Working";
            if (model.increase_rightshoulder_pain_Rotation == "true")
                cc_r_shoulder_increase = cc_r_shoulder_increase + ",Rotation";
            if (model.increase_rightshoulder_pain_Overhead_activities == "true")
                cc_r_shoulder_increase = cc_r_shoulder_increase + ",Overhead activities";


            if (!string.IsNullOrEmpty(cc_r_shoulder_increase))
                model.cc_r_shoulder = model.cc_r_shoulder + " Pain increases by " + cc_r_shoulder_increase.TrimStart(',') + ".";


            if (!string.IsNullOrEmpty(model.describe_leftknee))
                cc_l_knee = "The patient complains of left knee pain that is " + model.describe_leftknee + "/10, with 10 being the worst , which is ";


            if (model.txt_describe_leftknee_Constant == "true")
                cc_l_knee = cc_l_knee + ",Constant";
            if (model.txt_describe_leftknee_Intermittent == "true")
                cc_l_knee = cc_l_knee + ",Intermittent";
            if (model.txt_describe_leftknee_Sharp == "true")
                cc_l_knee = cc_l_knee + ",Sharp";
            if (model.txt_describe_leftknee_Electric == "true")
                cc_l_knee = cc_l_knee + ",Electric";
            if (model.txt_describe_leftknee_Shooting == "true")
                cc_l_knee = cc_l_knee + ",Shooting";
            if (model.txt_describe_leftknee_Throbbing == "true")
                cc_l_knee = cc_l_knee + ",Throbbing";
            if (model.txt_describe_leftknee_Pulsating == "true")
                cc_l_knee = cc_l_knee + ",Pulsating";
            if (model.txt_describe_leftknee_Dull == "true")
                cc_l_knee = cc_l_knee + ",Dull";
            if (model.txt_describe_leftknee_Achy == "true")
                cc_l_knee = cc_l_knee + ",Achy";

            model.cc_l_knee = cc_l_knee.TrimStart(',');

            if (model.increase_leftknee_pain_Squatting == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",Squatting";
            if (model.increase_leftknee_pain_Walking == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",Walking";
            if (model.increase_leftknee_pain_Climb == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",Climb stairs";
            if (model.increase_leftknee_pain_goingdown_stairs == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",going down stairs";
            if (model.increase_leftknee_pain_Standing == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",Standing";
            if (model.increase_leftknee_pain_getupfrom_chair == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",get up from chair";
            if (model.increase_leftknee_pain_getoutof_car == "true")
                cc_l_knee_increase = cc_l_knee_increase + ",get out of car";



            if (!string.IsNullOrEmpty(cc_l_knee_increase))
                model.cc_l_knee = model.cc_l_knee + " Pain increases by " + cc_l_knee_increase.TrimStart(',') + ".";


            if (!string.IsNullOrEmpty(model.describe_rightknee))
                cc_l_knee = "The patient complains of right right pain that is " + model.describe_rightknee + "/10, with 10 being the worst , which is ";


            if (model.txt_describe_rightknee_Constant == "true")
                cc_r_knee = cc_r_knee + ",Constant";
            if (model.txt_describe_rightknee_Intermittent == "true")
                cc_r_knee = cc_r_knee + ",Intermittent";
            if (model.txt_describe_rightknee_Sharp == "true")
                cc_r_knee = cc_r_knee + ",Sharp";
            if (model.txt_describe_rightknee_Electric == "true")
                cc_r_knee = cc_r_knee + ",Electric";
            if (model.txt_describe_rightknee_Shooting == "true")
                cc_r_knee = cc_r_knee + ",Shooting";
            if (model.txt_describe_rightknee_Throbbing == "true")
                cc_r_knee = cc_r_knee + ",Throbbing";
            if (model.txt_describe_rightknee_Pulsating == "true")
                cc_r_knee = cc_r_knee + ",Pulsating";
            if (model.txt_describe_rightknee_Dull == "true")
                cc_r_knee = cc_r_knee + ",Dull";
            if (model.txt_describe_rightknee_Achy == "true")
                cc_r_knee = cc_r_knee + ",Achy";

            model.cc_r_knee = cc_r_knee.TrimStart(',');

            if (model.increase_rightknee_pain_Squatting == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",Squatting";
            if (model.increase_rightknee_pain_Walking == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",Walking";
            if (model.increase_rightknee_pain_Climb == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",Climb stairs";
            if (model.increase_rightknee_pain_goingdownstairs == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",going down stairs";
            if (model.increase_rightknee_pain_Standing == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",Standing";
            if (model.increase_rightknee_pain_getupfromchair == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",get up from chair";
            if (model.increase_rightknee_pain_getoutofcar == "true")
                cc_r_knee_increase = cc_r_knee_increase + ",get out of car";

            if (!string.IsNullOrEmpty(cc_r_knee_increase))
                model.cc_r_knee = model.cc_r_knee + " Pain increases by " + cc_r_knee_increase.TrimStart(',') + ".";

            if (!string.IsNullOrEmpty(model.other_describe_part_value))
                cc_other_1 = "The patient complains of " + model.txt_other_describe_part + " pain that is " + model.other_describe_part_value + "/10, with 10 being the worst , which is ";


            if (model.txt_other_describe_part_Constant == "true")
                cc_other_1 = cc_other_1 + ",Constant";
            if (model.txt_other_describe_part_Intermittent == "true")
                cc_other_1 = cc_other_1 + ",Intermittent";
            if (model.txt_other_describe_part_Sharp == "true")
                cc_other_1 = cc_other_1 + ",Sharp";
            if (model.txt_other_describe_part_Electric == "true")
                cc_other_1 = cc_other_1 + ",Electric";
            if (model.txt_other_describe_part_Shooting == "true")
                cc_other_1 = cc_other_1 + ",Shooting";
            if (model.txt_other_describe_part_Throbbing == "true")
                cc_other_1 = cc_other_1 + ",Throbbing";
            if (model.txt_other_describe_part_Pulsating == "true")
                cc_other_1 = cc_other_1 + ",Pulsating";
            if (model.txt_other_describe_part_Dull == "true")
                cc_other_1 = cc_other_1 + ",Dull";
            if (model.txt_other_describe_part_Achy == "true")
                cc_other_1 = cc_other_1 + ",Achy";

            model.cc_other_1 = cc_other_1.TrimStart(',');


            if (!string.IsNullOrEmpty(model.other_describe_part_1_value))
                cc_other_2 = "The patient complains of " + model.txt_other_describe_part_1 + " pain that is " + model.other_describe_part_1_value + "/10, with 10 being the worst , which is ";


            if (model.txt_other_describe_part_1_Constant == "true")
                cc_other_2 = cc_other_2 + ",Constant";
            if (model.txt_other_describe_part_1_Intermittent == "true")
                cc_other_2 = cc_other_2 + ",Intermittent";
            if (model.txt_other_describe_part_1_Sharp == "true")
                cc_other_2 = cc_other_2 + ",Sharp";
            if (model.txt_other_describe_part_1_Electric == "true")
                cc_other_2 = cc_other_2 + ",Electric";
            if (model.txt_other_describe_part_1_Shooting == "true")
                cc_other_2 = cc_other_2 + ",Shooting";
            if (model.txt_other_describe_part_1_Throbbing == "true")
                cc_other_2 = cc_other_2 + ",Throbbing";
            if (model.txt_other_describe_part_1_Pulsating == "true")
                cc_other_2 = cc_other_2 + ",Pulsating";
            if (model.txt_other_describe_part_1_Dull == "true")
                cc_other_2 = cc_other_2 + ",Dull";
            if (model.txt_other_describe_part_1_Achy == "true")
                cc_other_2 = cc_other_2 + ",Achy";

            model.cc_other_2 = cc_other_2.TrimStart(',');

            //PE for Neck

            string pe_neck = "";


            service.Insert(model);

            return RedirectToAction("Create", "IntakeForm");
        }

        public IActionResult InitialIntake(int id = 0, string cid = "D0+uiKgkwrEBGsfnh/WxjL7DJy1b6slN3GKDydEumf0=")
        {
            var c_id = EncryptionHelper.Decrypt(cid);
            //int cmpid = Convert.ToInt32(c_id);
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

            var path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "IntakeForm//WR1.xml");
            ViewBag.WR1 = new SelectList(this.GetDropDownList(path, "WR1"), "Name", "Name");
            path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "IntakeForm//WR2.xml");
            ViewBag.WR2 = new SelectList(this.GetDropDownList(path, "WR2"), "Name", "Name");
            path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "IntakeForm//WR3.xml");
            ViewBag.WR3 = new SelectList(this.GetDropDownList(path, "WR3"), "Name", "Name");
            path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "IntakeForm//WR4.xml");
            ViewBag.WR4 = new SelectList(this.GetDropDownList(path, "WR4"), "Name", "Name");


            InitialIntake obj = new InitialIntake();

            if (id > 0)
            {
                obj = service.GetInitialIntakeById(id);
            }

            if (obj == null)
                obj = new InitialIntake();

            obj.cmp_id = cmpid.Value;
            return View(obj);
        }
        [HttpPost]
        public IActionResult InitialIntake(InitialIntake model)
        {

            var result = service.SaveInitialIntake(model);

            if (result != "0")
            {

                tbl_patient objPatient = new tbl_patient()
                {
                    account_no = null,
                    address = null,
                    city = null,
                    dob = model.dob,
                    email = null,
                    fname = model.fname,
                    gender = model.gender,
                    home_ph = null,
                    lname = model.lname,
                    mc = null,
                    mc_details = null,
                    mname = null,
                    mobile = null,
                    handeness = model.handedness,
                    ssn = null,
                    state = null,
                    age = 1,
                    cmp_id = model.cmp_id,
                };

                var patientId = _patientservices.Insert(objPatient);

                if (patientId > 0)
                {
                    var objIE = new tbl_patient_ie()
                    {

                        doa = model.doa,
                        doe = System.DateTime.Now,
                        location_id = model.location_id,
                        is_active = true,
                        patient_id = patientId,
                        intakeid = Convert.ToInt32(result)

                    };

                    var ie = _ieService.Insert(objIE);

                    if (ie > 0)
                    {
                        var objPage1 = new tbl_ie_page1()
                        {
                            pmh = this.getPMH(model),
                            psh = this.getPSH(model),
                            allergies = this.getAllergies(model),
                            ie_id = ie
                        };

                        _ieService.InsertPage1(objPage1);
                    }

                }
            }

            return RedirectToAction("Index", "Visit");
        }

        public IActionResult AIInitialIntake(int? locId, int? id, int? providerId)
        {
            if (providerId == null)
            {
                if (!HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).HasValue)
                {
                    //providerId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).Value;
                    providerId = Convert.ToInt32(HttpContext.Session.GetString("ProviderId") ?? "0");
                }
                else
                {
                    providerId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).Value;
                }
            }
            var templatePath = $"{Request.Scheme}://{Request.Host}/v2/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            //var templatePath = $"{Request.Scheme}://{Request.Host}/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            ViewBag.TemplateURL = templatePath + "/report-template.txt";
            ViewBag.TemplateDOCURL = templatePath + "/report-template-ie.docx";
            ViewBag.FormData = "";
            ViewBag.Id = "0";
            ViewBag.LocId = locId;

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            tbl_locations objLoc = new tbl_locations()
            {
                id = locId
            };
            var loc = _locservices.GetOne(objLoc);

            ViewBag.LocName = loc?.location;
            ViewBag.ProviderId = providerId;

            var locdata = _commonservices.GetLocations(Convert.ToInt32(cmpid));

            List<SelectListItem> lst = new List<SelectListItem>();

            int defaultlocation = HttpContext.Session.GetInt32(SessionKeys.SessionLocationId).Value;

            foreach (var item in locdata)
            {
                var obj = new SelectListItem()
                {
                    Text = item.Text,
                    Value = item.Value,
                    Selected = item.Value == locId.ToString() ? true : false
                };
                lst.Add(obj);

            }
            ViewBag.locList = lst;
            var providers = _userService.GetProviders(cmpid.Value);
            List<SelectListItem> lstp = new List<SelectListItem>();
            //int providerid = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId).Value;
            foreach (var item in providers)
            {
                var obj = new SelectListItem()
                {
                    Text = item.Text,
                    Value = item.Value,
                    Selected = item.Value == providerId.ToString() ? true : false
                };
                lstp.Add(obj);

            }
            ViewBag.providerList = lstp;



            //ViewBag.ProviderName = providerName == "--Select Provider--" ? "" : providerName;


            ViewBag.CmpId = cmpid.ToString();

            var _dataTreatment = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);
            ViewBag.Treatment = _dataTreatment;

            if (id != null)
            {
                var data = service.GetInitialIntakeAIById(id.Value);

                if (data != null)
                {
                    ViewBag.FormData = data.FormData;
                    ViewBag.Id = id;
                    ViewBag.SubmitDate = data.PatientSubmitDate;
                    ViewBag.Diagnosis = data.Diagnosis;
                    ViewBag.TreatmentHTML = data.Treatment;

                }
            }

            var client_code = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            ViewBag.UEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionUEBodyPart);
            ViewBag.LEBodyPart = HttpContext.Session.GetString(SessionKeys.SessionLEBodyPart);

            if (client_code.ToLower() == "qmppc")
                return PartialView("_IntakeQMPPC");
            else if (client_code.ToLower() == "bhfpc")
                return PartialView("_IntakeBHF");
            else if (client_code.ToLower() == "hposm")
                return PartialView("_IntakeHPOSM");
            else if (client_code.ToLower() == "imnpfhpc")
                return PartialView("_IntakeIMNPFHPC");
            else return PartialView("_IntakeBHF");
            //return View();
        }

        public IActionResult AIInitialIntakeV2(int? locId, int? id)
        {
            var templatePath = $"{Request.Scheme}://{Request.Host}/v2/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            //var templatePath = $"{Request.Scheme}://{Request.Host}/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
            ViewBag.TemplateURL = templatePath + "/report-template.txt";
            ViewBag.TemplateDOCURL = templatePath + "/report-template-ie.docx";
            ViewBag.FormData = "";
            ViewBag.Id = "0";
            ViewBag.LocId = locId;

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            tbl_locations objLoc = new tbl_locations()
            {
                id = locId
            };
            var loc = _locservices.GetOne(objLoc);

            ViewBag.LocName = loc?.location;
            ViewBag.CmpId = cmpid.ToString();

            var _dataTreatment = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);
            ViewBag.Treatment = _dataTreatment;

            if (id != null)
            {
                var data = service.GetInitialIntakeAIById(id.Value);

                if (data != null)
                {
                    ViewBag.FormData = data.FormData;
                    ViewBag.Id = id;
                    ViewBag.SubmitDate = data.PatientSubmitDate;
                    ViewBag.Diagnosis = data.Diagnosis;
                    ViewBag.TreatmentDesc = data.TreatmentDesc;
                }
            }

            return View();
        }

        public IActionResult PatientInitialIntake(string locId, string id)
        {
            var isSubmited = service.IsInitialIntakeAISubmited(id);

            if (!isSubmited)
            {
                //locId = "IUxzGfqagay0u+1EGXR1sf3dyOWvUoTf7VsQ86WOKws=";
                //id = "Vp3XDtf29OU9jCI/G7/j+VkAcoyR7jpPowp+SKKaa3M=";
                //var _locid = EncryptionHelper.Decrypt(locId);
                //var _id = EncryptionHelper.Decrypt(id);
                var _locid = locId;
                var _id = id;
                var templatePath = $"{Request.Scheme}://{Request.Host}/v2/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
                //var templatePath = $"{Request.Scheme}://{Request.Host}/ReportTemplate/" + HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
                ViewBag.TemplateURL = templatePath + "/report-template.txt";
                ViewBag.TemplateDOCURL = templatePath + "/report-template-ie.docx";
                ViewBag.FormData = "";
                ViewBag.Id = "0";
                ViewBag.LocId = _locid;
                ViewBag.SubmitDate = System.DateTime.Now;

                tbl_locations objLoc = new tbl_locations()
                {
                    id = Convert.ToInt32(_locid)
                };
                var loc = _locservices.GetOne(objLoc);

                ViewBag.LocName = loc?.location;

                if (_id != null)
                {
                    var data = service.GetInitialIntakeAIById(Convert.ToInt32(_id));

                    if (data != null)
                    {
                        // ViewBag.FormData = data.FormData;
                        ViewBag.FormData = data.FormData;
                        ViewBag.Id = _id;

                    }
                }
                ViewBag.IsSubmited = false;
            }
            else
            {
                ViewBag.IsSubmited = true;
            }

            return View();
        }

        private List<IntakeDropDown> GetDropDownList(string path, string node)
        {
            var intakeDropDowns = new List<IntakeDropDown>();

            XDocument doc = XDocument.Load(path);

            foreach (var item in doc.Descendants(node))
            {
                intakeDropDowns.Add(new IntakeDropDown
                {
                    Name = item.Element("Name")?.Value
                });
            }

            return intakeDropDowns;
        }

        [HttpPost]
        public IActionResult SaveForm([FromBody] object formData)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(formData);
            var model = new AIIntakeFormModel();
            try
            {
                model = System.Text.Json.JsonSerializer.Deserialize<AIIntakeFormModel>(json);
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Path: {ex.Path}");
                Console.WriteLine(ex.Message);
            }

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var result = "0";

            DateTime doa = Convert.ToDateTime(model.DOA);

            if (IsPatientPresent(model.FN, model.LN, doa) && model.Id == "0")
            {
                return Json(new { success = false, message = "The patient already exists in the system. Please verify the patient details before proceeding." });
            }

            if (model != null)
            {
                InitialIntakeAI initialIntakeAI = new InitialIntakeAI()
                {
                    Id = model.Id == "" ? 0 : Convert.ToInt32(model.Id),
                    CmpId = cmpid,
                    Visit_Type = "IE",
                    DOA = DateTime.TryParse(model.DOA, out var parsedDOA) ? parsedDOA : (DateTime?)null,
                    DOB = DateTime.TryParse(model.DOB, out var parsedDOB) ? parsedDOB : (DateTime?)null,
                    //DOE = System.DateTime.Now,
                    DOE = DateTime.TryParse(model.DOE, out var parsedDOE) ? parsedDOE : (DateTime?)null,
                    FormData = json,
                    FN = model.FN,
                    LN = model.LN,
                    PatientSubmitDate = DateTime.TryParse(model.PatientSubmitDate, out var PatientSubmitDate) ? parsedDOA : (DateTime?)null,
                    LocationId = string.IsNullOrEmpty(model.LocationId) ? null : Convert.ToInt32(model.LocationId),
                    DLPath = model.DLPath,
                    Diagnosis = model.Diagnosis,
                    Treatment = model.Treatment,
                    TreatmentIds = model.TreatmentIds,
                    TreatmentDelimitDesc = model.TreatmentDelimitDesc
                };
                result = service.SaveInitialIntakeAI(initialIntakeAI);

                var InjuryType = "MM";


                if (model.InjuryType == "work-related")
                    InjuryType = "WC";
                else if (model.InjuryType == "lien")
                    InjuryType = "Lien";
                else
                    InjuryType = model.InjuryType;


                if (initialIntakeAI.Id == 0)
                {
                    if (result != "0")
                    {
                        tbl_patient objPatient = new tbl_patient()
                        {
                            account_no = null,
                            address = null,
                            city = null,
                            dob = string.IsNullOrEmpty(model.DOB) ? null : Convert.ToDateTime(model.DOB),
                            email = null,
                            fname = model.FN,
                            gender = model.Gender.ToLower() == "male" ? "1" : "2",
                            home_ph = null,
                            lname = model.LN,
                            mc = null,
                            mc_details = null,
                            mname = null,
                            mobile = null,
                            handeness = model.DominantHand,
                            ssn = null,
                            state = null,
                            age = string.IsNullOrEmpty(model.Age) ? 0 : Convert.ToInt16(model.Age),
                            cmp_id = cmpid,
                            
                        };

                        var patientId = _patientservices.Insert(objPatient);

                        if (patientId > 0)
                        {

                            var objIE = new tbl_patient_ie()
                            {

                                doa = string.IsNullOrEmpty(model.DOA) ? null : Convert.ToDateTime(model.DOA),
                                doe = string.IsNullOrEmpty(model.DOE) ? null : Convert.ToDateTime(model.DOE),
                                location_id = string.IsNullOrEmpty(model.LocationId) ? 0 : Convert.ToInt16(model.LocationId),
                                is_active = true,
                                patient_id = patientId,
                                compensation = InjuryType,
                                accident_type= model.AccidentType,
                                provider_id = string.IsNullOrEmpty(model.ProviderId) ? null : Convert.ToInt32(model.ProviderId),
                                intakeid = Convert.ToInt32(result)

                            };

                            var ie = _ieService.Insert(objIE);

                            if (ie > 0)
                            {
                                var defaultPage1 = _defaultSettingService.GetOnePage1(cmpid.Value);
                                string history = "";
                                if (defaultPage1 != null)
                                {

                                    history = defaultPage1.history;

                                    history = history.Replace("#fname", model.FN);
                                    history = history.Replace("#fn", model.FN);
                                    history = history.Replace("#lname", model.LN);
                                    history = history.Replace("#ln", model.LN);
                                   
                                    history = history.Replace("#dos", model.DOE);
                                    //history = history.Replace("#doi", model.DOA);
                                    history = history.Replace("#handedness", model.DominantHand);
                                    history = history.Replace("#sex", model.Gender);
                                    history = history.Replace("#gender", Common.GetMrMrsFromSex(model.Gender == "male" ? "1" : "2"));

                                }
                                string assessment = defaultPage1.daignosis_desc == null ? "" : defaultPage1.daignosis_desc;
                                assessment = assessment.Replace("#sex", model.Gender);

                                var vital = "";

                                if (!string.IsNullOrEmpty(model.Height) && !string.IsNullOrEmpty(model.Weight))
                                {
                                    vital = "The patient’s height is " + model.Height + ", weight is " + model.Weight + " pounds, and BMI is _____.";
                                }

                                var objPage1 = new tbl_ie_page1()
                                {
                                    pmh = string.IsNullOrEmpty(string.Join(", ", model.PMH)) ? defaultPage1.pmh : string.Join(", ", model.PMH),
                                    dd = defaultPage1.dd,
                                    psh = string.IsNullOrEmpty(string.Join(", ", model.PSH)) ? defaultPage1.psh : string.Join(", ", model.PSH),
                                    bodypart = string.Join(",", model.Complaints),
                                    allergies = defaultPage1.allergies,
                                    assessment = string.IsNullOrEmpty(model.Diagnosis) ? assessment : model.Diagnosis,
                                    ie_id = ie,
                                    vital = string.IsNullOrEmpty(vital) ? defaultPage1.vital : vital,
                                    cc = string.IsNullOrEmpty(this.GetCC(model)) ? defaultPage1.cc : this.GetCC(model),
                                    daignosis_desc = assessment,
                                    pe = defaultPage1.pe,
                                    family_history = defaultPage1.family_history,
                                    history = history,
                                    medication = defaultPage1.medication,
                                    note = defaultPage1.note,
                                    occupation = defaultPage1.occupation,
                                    plan = defaultPage1.plan,
                                    rom = defaultPage1.rom,
                                    social_history = defaultPage1.social_history,
                                    work_status = defaultPage1.work_status
                                };

                                _ieService.InsertPage1(objPage1);

                                var defaultPage2 = _defaultSettingService.GetOnePage2(cmpid.Value);

                                if (defaultPage2 != null)
                                {
                                    var objPage2 = new tbl_ie_page2()
                                    {
                                        ie_id = ie,
                                        ros = defaultPage2?.ros,
                                        aod = defaultPage2?.aod,
                                        other = defaultPage2?.other,
                                        cmp_id = cmpid.Value,
                                        patient_id = patientId,
                                    };

                                    _ieService.InsertPage2(objPage2);
                                }

                                var defaultNE = _defaultSettingService.GetOneNE(cmpid.Value);

                                if (defaultNE != null)
                                {
                                    var objNE = new tbl_ie_ne()
                                    {
                                        ie_id = ie,
                                        neurological_exam = defaultNE?.neurological_exam,
                                        sensory = defaultNE?.sensory,
                                        manual_muscle_strength_testing = defaultNE?.manual_muscle_strength_testing,
                                        other_content = defaultNE?.other_content,
                                        cmp_id = cmpid.Value,
                                        patient_id = patientId,
                                    };

                                    _ieService.InsertNE(objNE);
                                }

                                var objOther = new tbl_ie_other()
                                {
                                    ie_id = ie,
                                   
                                    patient_id = patientId,
                                    treatment_delimit = model.TreatmentIds,
                                    treatment_delimit_desc = model.TreatmentDelimitDesc,
                                    treatment_details = model.TreatmentDesc,
                                    followup_duration = HttpContext.Session.GetString(SessionKeys.SessionFUDate)
                                };

                                _ieService.InsertOtherPage(objOther);


                                var _page3Setting = _settingservices.GetOne(cmpid.Value);

                                var dataPage3 = new tbl_ie_page3();

                                dataPage3.diagcervialbulge_study = "1";
                                dataPage3.diagcervialbulge_comma = _page3Setting.diagcervialbulge_comma;
                                dataPage3.diagthoracicbulge_study = "1";
                                dataPage3.diagthoracicbulge_comma = _page3Setting.diagthoracicbulge_comma;
                                dataPage3.diaglumberbulge_study = "1";
                                dataPage3.diaglumberbulge_comma = _page3Setting.diaglumberbulge_comma;
                                dataPage3.diagleftshoulder_study = "1";
                                dataPage3.diagleftshoulder_comma = _page3Setting.diagleftshoulder_comma;
                                dataPage3.diagrightshoulder_study = "1";
                                dataPage3.diagrightshoulder_comma = _page3Setting.diagrightshoulder_comma;
                                dataPage3.diagleftknee_study = "1";
                                dataPage3.diagleftknee_comma = _page3Setting.diagleftknee_comma;
                                dataPage3.diagrightknee_study = "1";
                                dataPage3.diagrightknee_comma = _page3Setting.diagrightknee_comma;
                                dataPage3.diaglumberbulge_study = "1";
                                dataPage3.diaglumberbulge_comma = _page3Setting.diaglumberbulge_comma;

                                dataPage3.other1_study = "0";
                                dataPage3.other1_comma = _page3Setting.other1_comma;
                                dataPage3.other2_study = "0";
                                dataPage3.other2_comma = _page3Setting.other2_comma;
                                dataPage3.other3_study = "0";
                                dataPage3.other3_comma = _page3Setting.other3_comma;
                                dataPage3.other4_study = "0";
                                dataPage3.other4_comma = _page3Setting.other4_comma;
                                dataPage3.other5_study = "0";
                                dataPage3.other5_comma = _page3Setting.other5_comma;
                                dataPage3.other6_study = "0";
                                dataPage3.other6_comma = _page3Setting.other6_comma;
                                dataPage3.other7_study = "0";
                                dataPage3.other7_comma = _page3Setting.other7_comma;

                                dataPage3.gait = string.IsNullOrEmpty(dataPage3.gait) ? _page3Setting.gait_default : dataPage3.gait;


                                var _defaultdata = _defaultService.GetOneByCompany(cmpid.Value);

                                if (_defaultdata != null)
                                {
                                    dataPage3.goal = _defaultdata.goal;
                                    dataPage3.care = _defaultdata.care;
                                }
                                dataPage3.ie_id= ie;
                                dataPage3.cmp_id = cmpid.Value;

                                _ieService.InsertPage3(dataPage3);

                                using JsonDocument doc = JsonDocument.Parse(json);

                                string[] planUTPI = Array.Empty<string>();

                                // 1. Check if the property exists
                                if (doc.RootElement.TryGetProperty("PlanUTPI", out JsonElement planElement))
                                {
                                    // 2. Optional: Ensure it is actually an array before enumerating
                                    if (planElement.ValueKind == JsonValueKind.Array)
                                    {
                                        planUTPI = planElement
                                            .EnumerateArray()
                                            .Select(x => x.GetString() ?? string.Empty)
                                            .ToArray();
                                    }
                                }
                                foreach (string data in planUTPI)
                                {
                                    var _obj = new ProcedureDetailsIntakeVM()
                                    {
                                        PatientIEID = ie,
                                        MCode = data,
                                        PatientFuID = null,
                                        Cmp_Id = cmpid.Value,
                                        Date = DateTime.TryParse(model.DOE, out var pDOE) ? pDOE : (DateTime?)null,
                                        IsExecuted = true
                                    };
                                    _pocservices.SaveProcedureDetailsIntake(_obj);
                                }

                                string[] recommendation = Array.Empty<string>();

                                // 1. Check if the property exists
                                if (doc.RootElement.TryGetProperty("Recommendation", out JsonElement recommendationElement))
                                {
                                    // 2. Optional: Ensure it is actually an array before enumerating
                                    if (recommendationElement.ValueKind == JsonValueKind.Array)
                                    {
                                        recommendation = recommendationElement
                                            .EnumerateArray()
                                            .Select(x => x.GetString() ?? string.Empty)
                                            .ToArray();
                                    }
                                }
                                foreach (string data in recommendation)
                                {
                                    var _obj = new ProcedureDetailsIntakeVM()
                                    {
                                        PatientIEID = ie,
                                        MCode = data,
                                        PatientFuID = null,
                                        Cmp_Id = cmpid.Value,
                                        Date = DateTime.TryParse(model.DOE, out var pDOE) ? pDOE : (DateTime?)null,
                                        IsExecuted = false
                                    };
                                    _pocservices.SaveProcedureDetailsIntake(_obj);
                                }
                            }
                        }
                    }
                }
                else
                {
                    var objIE = new tbl_patient_ie()
                    {
                        doa = string.IsNullOrEmpty(model.DOA) ? null : Convert.ToDateTime(model.DOA),
                        doe = string.IsNullOrEmpty(model.DOE) ? null : Convert.ToDateTime(model.DOE),
                        compensation = InjuryType,
                        intakeid = initialIntakeAI.Id
                    };
                    _ieService.UpdateFromIntake(objIE);

                    var objPage1 = new tbl_ie_page1()
                    {

                        bodypart = string.Join(",", model.Complaints),
                        id = initialIntakeAI.Id

                    };

                    _ieService.UpdatePage1Intake(objPage1);
                }

                //return RedirectToAction("Index", "Visit");
            }
            return Json(new { success = true, message = "Intake form summited successfully.", id = result, locid = model.LocationId });
        }

        [HttpPost]
        public IActionResult SavePatientForm([FromBody] object formData)
        {
            try
            {
                var json = System.Text.Json.JsonSerializer.Serialize(formData);
                var model = System.Text.Json.JsonSerializer.Deserialize<AIIntakeFormModel>(json);
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                if (model != null)
                {
                    InitialIntakeAI initialIntakeAI = new InitialIntakeAI()
                    {
                        Id = model.Id == "" ? 0 : Convert.ToInt32(model.Id),
                        CmpId = cmpid,
                        Visit_Type = "IE",
                        DOA = DateTime.TryParse(model.DOA, out var parsedDOA) ? parsedDOA : (DateTime?)null,
                        DOB = DateTime.TryParse(model.DOB, out var parsedDOB) ? parsedDOB : (DateTime?)null,
                        // DOE = System.DateTime.Now,
                        FormData = json,
                        FN = model.FN,
                        LN = model.LN,
                        PatientSubmitDate = System.DateTime.Now,
                        LocationId = string.IsNullOrEmpty(model.LocationId) ? null : Convert.ToInt32(model.LocationId),
                        DLPath = model.DLPath
                    };
                    var result = service.SaveInitialIntakeAI(initialIntakeAI);
                }
                return Json(new { success = true, message = "Intake form summited successfully." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false });
            }
        }

        [HttpPost]
        public IActionResult ExportWord(string htmlContent)
        {

            string filePath = "", docName = "", patientName = "", injDocName = "", dos = "", dob = "";
            string[] splitContent;


            // Create a new DOCX package
            using (MemoryStream memStream = new MemoryStream())
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
                {

                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    var headerPart = mainPart.AddNewPart<HeaderPart>();
                    var footerPart = mainPart.AddNewPart<FooterPart>();

                    // Create the main document part content
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());


                    // Define the font and size
                    RunProperties runProperties = new RunProperties();
                    RunFonts runFonts = new RunFonts() { Ascii = "Times New Roman" };
                    FontSize fontSize = new FontSize() { Val = "24" }; // Font size 12 (in half-point format)

                    // Apply the font settings to the RunProperties
                    runProperties.Append(runFonts);
                    runProperties.Append(fontSize);

                    // Parse the HTML content and append it to the document
                    HtmlConverter converter = new HtmlConverter(mainPart);
                    // Clean invalid or empty <img> tags
                    htmlContent = Regex.Replace(htmlContent, @"<img[^>]*base64,\s*""[^>]*>", "", RegexOptions.IgnoreCase);

                    IList<OpenXmlCompositeElement> generatedBody = converter.Parse(htmlContent);

                    // Iterate over the parsed elements and apply RunProperties
                    foreach (var element in generatedBody)
                    {
                        foreach (var run in element.Descendants<Run>()) // Find all Run elements in the element
                        {
                            run.PrependChild(runProperties.CloneNode(true)); // Apply the font properties to each Run
                        }

                        // Append each element to the body
                        body.Append(element.CloneNode(true));
                    }


                    var header = new Header(new Paragraph(new Run(new Text("Header Test"))));
                    HeaderReference headerReference = new HeaderReference() { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) };

                    headerPart.Header = header;

                    mainPart.Document.Body.Append(new SectionProperties(headerReference));

                }
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

                string subPath = "Report/" + cmpid; // Your code goes here

                bool exists = System.IO.Directory.Exists(subPath);

                if (!exists)
                    System.IO.Directory.CreateDirectory(subPath);

                filePath = subPath + "/IE.docx";

                // Save the memory stream to a file
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                {
                    memStream.WriteTo(fileStream);
                }

                try
                {
                    string filepathFrom = Path.Combine(Environment.WebRootPath, "Uploads/HeaderTemplate") + "//" + HttpContext.Session.GetString(SessionKeys.SessionHeaderTemplate); ;


                    string filepathTo = filePath;
                    AddHeaderFromTo(filepathFrom, filepathTo, patientName, dos, "");
                    //if (DoesFooterExist(filepathFrom))
                    //    AddFooterFromTo(filepathFrom, filepathTo, patientName, dos, dob);
                }
                catch (Exception ex)
                {
                }

                //filePath = _env.ContentRootPath + filePath;
                memStream.Position = 0; // VERY IMPORTANT
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(
                    //memStream.ToArray(),
                    fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    docName
                );
            }


        }

        [HttpPost]
        public async Task<IActionResult> UploadAudio(IFormFile file)
        {
            var path = Path.Combine("wwwroot/audio", file.FileName);

            using (var stream = new FileStream(path, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return Ok();
        }

        [HttpPost]
        public async Task<IActionResult> UploadDL(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded");

            // Save file temporarily
            var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Uploads/DriverLicense");

            if (!Directory.Exists(uploadsFolder))
                Directory.CreateDirectory(uploadsFolder);

            var fileName = Guid.NewGuid().ToString("N").Substring(0, 10) + "_" + file.FileName;
            var filePath = Path.Combine(uploadsFolder, fileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // OCR Processing
            var text = await _azureService.ExtractTextFromImage(filePath);
            var data = _azureService.ParseData(text);

            data.DOB = string.IsNullOrEmpty(data.DOB) ? "" : data.DOB.Replace("-", "/");
            data.Gender = string.IsNullOrEmpty(data.Gender) ? "male" : data.Gender;


            data.FileName = "wwwroot/Uploads/DriverLicense/" + fileName;
            return Json(new
            {
                success = true,
                extractedText = text,
                parsedData = data,
            });
        }

        [HttpPost]
        public IActionResult GetDaignoCodeList(string bodyparts, int id)
        {
            try
            {

                var page1Data = service.GetInitialIntakeAIById(id);

                string assetment = "";

                if (page1Data != null)
                    assetment = page1Data.Diagnosis;

                bodyparts = bodyparts.Replace("_", " ");
                bodyparts = bodyparts.TrimEnd();
                ViewBag.BodyPart = bodyparts.ToUpper();
                var _bodyparts = _commonservices.GetBodyPartIntake(bodyparts);
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

                var formatted = string.Join("','", _bodyparts.Split(',').Select(x => x.Trim()));

                var parts = formatted.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                string result = parts.Length > 1 ? parts[1] : parts[0];

                string cnd = " and d.cmp_id=" + cmpid + " and (d.BodyPart like '%" + bodyparts + "%' or d.Description like '%" + _bodyparts + "%') order by d.display_order ASC";

                var data = _diagcodesService.GetAllWithGroups(cnd);

                var cmpIdInt = Convert.ToInt32(cmpid);

                var datavm = (from c in data
                              select new DaignoCodeVM
                              {
                                  DaignoCodeId = c.Id.Value,
                                  Description = c.Description,
                                  GroupDisplayOrder = c.GroupDisplayOrder.HasValue ? c.GroupDisplayOrder.Value : 100,
                                  DiagCode = c.DiagCode,
                                  DiagCodeGroup = c.DiagCodeGroup == null ? "0" : c.DiagCodeGroup.ToString(),
                                  GroupName = c.GroupName == null ? "NA" : c.GroupName.ToString(),
                                  IsSelect = assetment != null ? (assetment.IndexOf(c.DiagCode) > 0 ? true : c.PreSelect) : c.PreSelect,
                                  Display_Order = c.display_order,
                                  cmp_id = c.cmp_id

                              }).ToList().Where(x => x.cmp_id == cmpIdInt).OrderBy(x => x.GroupDisplayOrder);


                var groupedData = datavm
                .GroupBy(x => x.GroupName)
                .OrderBy(g => g.First().GroupDisplayOrder)
                .ToList();

                return PartialView("_DaignoCode", groupedData);


            }
            catch (Exception ex)
            {
                SaveLog(ex, "GetDaignoCodeList");
            }

            return View();
        }

        public IActionResult TodayVisitList(int locid)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var model = service.GetInitialIntakeByLocationId(locid, cmpid.Value);
            return PartialView("_TodayVisitList", model);
        }

        [HttpPost]
        public IActionResult CheckPatientNameOnly(VisitVM model)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                string cnd = "";

                if (!string.IsNullOrEmpty(model.mname))
                    cnd = " and fname='" + model.fname + "' and lname='" + model.lname + "' and mname='" + model.mname + "'  and cmp_id=" + cmpid;
                else
                    cnd = " and fname='" + model.fname + "' and lname='" + model.lname + "' and mname is null  and cmp_id=" + cmpid;

                if (model.doa.HasValue == false)
                    cnd = cnd + " and doa is null";
                else
                    cnd = cnd + " and doa='" + model.doa.Value.ToString("yyyy-MM-dd") + "'";

                if (!string.IsNullOrEmpty(model.account_no))
                    cnd = cnd + " and account_no='" + model.account_no + "'";


                var data = _patientservices.GetAll(cnd);

                if (data.Count > 0)
                    return Json(new { result = 1, data = data });
                else
                    return Json(new { result = 0 });
            }
            catch (Exception ex)
            {
                SaveLog(ex, "CheckPatientNameOnly");
                return Json(new { result = -1, message = "An error occurred " });
            }
        }

        #region private method
        public void AddHeaderFromTo(string filepathFrom, string filepathTo, string patientName = "", string dos = "", string provName = "")
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument
                wdDoc = WordprocessingDocument.Open(filepathTo, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                // Delete the existing header part.
                mainPart.DeleteParts(mainPart.HeaderParts);

                // Create a new header part.
                DocumentFormat.OpenXml.Packaging.HeaderPart headerPart =
            mainPart.AddNewPart<HeaderPart>();

                // Get Id of the headerPart.
                string rId = mainPart.GetIdOfPart(headerPart);

                // Feed target headerPart with source headerPart.
                using (WordprocessingDocument wdDocSource =
                    WordprocessingDocument.Open(filepathFrom, true))
                {
                    DocumentFormat.OpenXml.Packaging.HeaderPart firstHeader =
            wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

                    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

                    if (firstHeader != null)
                    {

                        headerPart.FeedData(firstHeader.GetStream());
                    }

                    // Copy Image Parts
                    //foreach (var imagePart in firstHeader.ImageParts)
                    //{
                    //    // Add image part to the target header
                    //    ImagePart newImagePart = headerPart.AddImagePart(imagePart.ContentType);

                    //    // Copy image stream
                    //    using (Stream imageStream = imagePart.GetStream())
                    //    {
                    //        newImagePart.FeedData(imageStream);
                    //    }
                    //}

                    // Copy Image Parts
                    Dictionary<string, string> imageRelMapping = new Dictionary<string, string>();

                    foreach (var imagePart in firstHeader.ImageParts)
                    {
                        // Add a new image part to the target header
                        ImagePart newImagePart = headerPart.AddImagePart(imagePart.ContentType);

                        // Copy the image data
                        using (Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            newImagePart.FeedData(imageStream);
                        }

                        // Map the old relationship ID to the new image part ID
                        string oldRelId = firstHeader.GetIdOfPart(imagePart);
                        string newRelId = headerPart.GetIdOfPart(newImagePart);
                        imageRelMapping[oldRelId] = newRelId;
                    }

                    // Update relationships in header XML
                    UpdateHeaderXml(headerPart, imageRelMapping, provName);



                    foreach (var para in headerPart.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                    {
                        // Normalize spacing
                        var pPr = para.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>().FirstOrDefault();
                        if (pPr == null)
                        {
                            pPr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                            para.PrependChild(pPr);
                        }

                        var spacing = pPr.Elements<DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines>().FirstOrDefault();
                        if (spacing == null)
                        {
                            spacing = new DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines();
                            pPr.Append(spacing);
                        }

                        spacing.Before = "0";         // Remove space before
                        spacing.After = "0";          // Remove space after
                        spacing.Line = "240";         // 240 = single line spacing
                        spacing.LineRule = DocumentFormat.OpenXml.Wordprocessing.LineSpacingRuleValues.Auto;
                    }

                    //Dictionary<string, string> textReplacements = new Dictionary<string, string>
                    //    {
                    //            { "@drname@", "Dr. Patel" }  // Replace with your dynamic name
                    //    };
                    //ReplacePlaceholdersInHeader(headerPart, textReplacements);
                }

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                var restheaderPart = mainPart.AddNewPart<HeaderPart>("Rest");
                restheaderPart.Header = CreateHeaderWithPageNumber(patientName, "");
                if (cmpid == 7 || cmpid == 13 || cmpid == 18)
                {
                    if (!string.IsNullOrEmpty(dos))
                    {
                        string _dos = Common.commonDate(Convert.ToDateTime(dos), HttpContext.Session.GetString(SessionKeys.SessionDateFormat));
                        if (cmpid == 18)
                        {
                            restheaderPart.Header = CreateHeaderWithPageNumber("Re: " + patientName, "");
                        }
                        else
                            restheaderPart.Header = CreateHeaderWithPageNumber(patientName, _dos);
                    }
                }
                else
                {

                    restheaderPart.Header = CreateHeaderWithPageNumber(patientName, "");
                }


                //  restheaderPart.Header = new Header(new Paragraph("Purav\nSandip"));
                string restId = mainPart.GetIdOfPart(restheaderPart);
                // Get SectionProperties and Replace HeaderReference with new Id.
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs =
            mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to headers.
                    sectPr.RemoveAllChildren<HeaderReference>();
                    sectPr.Append(new TitlePage());
                    // Create the new header reference node.
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.First, Id = rId });
                    if (cmpid == 7 || cmpid == 13 || cmpid == 18)
                        sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Default, Id = restId });
                }
            }
        }
        // Method to update header XML to reference new image relationships
        private static void UpdateHeaderXml(HeaderPart headerPart, Dictionary<string, string> imageRelMapping, string provName)
        {
            string headerXml;

            // Read the existing header XML
            using (StreamReader reader = new StreamReader(headerPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                headerXml = reader.ReadToEnd();
            }

            // Replace old relationship IDs with new ones
            foreach (var kvp in imageRelMapping)
            {
                headerXml = headerXml.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
            }
            //this 2 lines
            //headerXml = headerXml
            //         .Replace("<w:t>@</w:t><w:t>drname</w:t><w:t>@</w:t>", "<w:t>@drname@</w:t>");

            // Step 3: Replace placeholders
            headerXml = headerXml.Replace("drname", provName == null ? "" : provName);

            // Write the updated XML back to the header part
            using (MemoryStream memStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(memStream))
                {
                    writer.Write(headerXml);
                    writer.Flush();
                    memStream.Position = 0;
                    headerPart.FeedData(memStream);
                }
            }
        }
        public Header CreateHeaderWithPageNumber(string text1, string text2)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            string fontSize = HttpContext.Session.GetString(SessionKeys.SessionFontSize);
            string fontFamily = HttpContext.Session.GetString(SessionKeys.SessionFontFamily);

            fontFamily = !string.IsNullOrEmpty(fontFamily) ? fontFamily : "Times New Roman";
            fontSize = !string.IsNullOrEmpty(fontSize) ? (Convert.ToInt16(fontSize) * 2).ToString() : "20";

            if (text2 != "")
            {
                return new Header(
                    new Paragraph(
                        new Run(
                            new RunProperties(
                     new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
                ),
                            new Text(text1) // First line
                        ),
                        new Break(), // Line break
                        new Run(
                            new RunProperties(
                      new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
                ),
                            new Text(text2) // Second line
                        ),
                        new Break(), // Another line break if needed
                        new Run(
                            new RunProperties(
                     new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
                ),
                            new Text("Page ") // Static "Page " text
                        ),
        // Explicit space
        new Run(
            new Text(" ")
            {
                Space = SpaceProcessingModeValues.Preserve
            },
             new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
        ),
                        new Run(
                            new SimpleField() // Dynamic page number field
                            {
                                Instruction = "PAGE", // Specifies the field type
                            }
                        )
                    )
                );
            }
            else
            {
                if (cmpid == 18)
                {

                    return new Header(
                          new Paragraph(
                              new Run(
                                   new RunProperties(
                           new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                             new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
                       ),
                                  new Text(text1) // First line
                              ),
                               new Break(), // Line break

                              new Run(
                                   new RunProperties(
                                         new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                           new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
                                   ),
                                   new Text("Page ") // Static "Page " text
                              ),
        // Explicit space
        new Run(
            new Text(" ")
            {
                Space = SpaceProcessingModeValues.Preserve
            }
        ),
                              new Run(
                                   new RunProperties(
                             new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }

                       ),
                                  new SimpleField() // Dynamic page number field
                                  {
                                      Instruction = "PAGE", // Specifies the field type
                                  }
                              )
                          )
                      );
                }
                else
                {
                    return new Header(
                       new Paragraph(
                           new Run(
                                new RunProperties(
                         new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily }
                    ),
                               new Text(text1) // First line
                           ),
                           new Break(), // Line break

                           new Run(
                                new RunProperties(
                          new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }
                    ),
                               new Text("Page ") // Static "Page " text
                           ),
        // Explicit space
        new Run(
            new Text(" ")
            {
                Space = SpaceProcessingModeValues.Preserve
            }
        ),
                           new Run(
                                new RunProperties(
                         new RunFonts { Ascii = fontFamily, HighAnsi = fontFamily },
                          new FontSize { Val = fontSize },
                           new FontSizeComplexScript { Val = fontSize }

                    ),
                               new SimpleField() // Dynamic page number field
                               {
                                   Instruction = "PAGE", // Specifies the field type
                               }
                           )
                       )
                   );
                }
            }
        }
        private string getPMH(InitialIntake model)
        {
            string str = "Noncontributory.";

            if (model.PMHNone == "false")
            {
                str = "";

                if (model.PMHDiabetes == "true")
                    str = ", Diabetes";
                if (model.PMHHTN == "true")
                    str += ", HTN";
                if (model.PMHHLD == "true")
                    str += ", HLD";
                if (model.PMHAsthma == "true")
                    str += ", Asthma";
                if (model.PMHCardiac == "true")
                    str += ", Cardiac";
                if (model.PMHThyroid == "true")
                    str += ", Thyroid";
                if (model.PMHCA == "true")
                    str += ", CA";

                if (!string.IsNullOrEmpty(str))
                {
                    str = " Patient have " + str.TrimStart(',');
                }
            }

            return str;
        }
        private string getPSH(InitialIntake model)
        {
            string str = "Noncontributory.";

            if (model.PSHNone == "false")
            {
                str = "";

                if (!string.IsNullOrEmpty(model.description_of_the_PSH))
                    str = model.description_of_the_PSH;
            }

            return str;
        }
        private string getAllergies(InitialIntake model)
        {
            string str = "NO KNOWN DRUG ALLERGIES.";

            if (model.DrugAllergy_yes.ToLower() == "yes")
            {
                str = "";

                if (!string.IsNullOrEmpty(model.DrugAllergy))
                    str = model.DrugAllergy;
            }

            return str;
        }
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
                Message = msg
            };
            new LogService().Insert(logdata);
        }
        private string GetCC(AIIntakeFormModel model)
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            if (cmpid == "21")
            {
                return "";
            }
            else
            {
                string cc_rsh = "", cc_rsh_difficulty = "", cc_rsh_imporve = "",
                    cc_lsh = "", cc_lsh_difficulty = "", cc_lsh_imporve = "",
                      cc_lkn = "", cc_lkn_difficulty = "", cc_lkn_imporve = "",
                       cc_rkn = "", cc_rkn_difficulty = "", cc_rkn_imporve = "";

                //right soulder
                if (!string.IsNullOrEmpty(model.RShPain))
                    cc_rsh = "The patient’s right shoulder pain level is " + model.RShPain + "/10. ";
                if (model.RShSymptoms?.Count > 0)
                    cc_rsh = cc_rsh + "The patient complains of " + string.Join(", ", model.RShSymptoms) + ". ";

                if (model.RShReachOverhead?.ToLower() == "yes")
                    cc_rsh_difficulty = "Overhead";
                if (model.RShReachBack?.ToLower() == "yes")
                    cc_rsh_difficulty = cc_rsh_difficulty + ", Back";
                if (model.RShSleepIssue?.ToLower() == "yes")
                    cc_rsh_difficulty = cc_rsh_difficulty + ", Sleeping";

                if (!string.IsNullOrEmpty(cc_rsh_difficulty))
                    cc_rsh = cc_rsh + "The patient has difficulty " + cc_rsh_difficulty + " on the right shoulder. ";

                if (model.RShImprove?.Count > 0)
                    cc_rsh = cc_rsh + "There has been improvement with " + string.Join(", ", model.RShImprove) + ".";
                //else
                //    cc_rsh = cc_rsh + "There has been no improvement with physical therapy.";

                //left soulder
                if (!string.IsNullOrEmpty(model.LShPain))
                    cc_lsh = "The patient’s left shoulder pain level is " + model.LShPain + "/10. ";
                if (model.LShSymptoms?.Count > 0)
                    cc_lsh = cc_lsh + "The patient complains of " + string.Join(", ", model.LShSymptoms) + ". ";


                if (model.LShReachOverhead?.ToLower() == "yes")
                    cc_lsh_difficulty = "Overhead";
                if (model.LShReachBack?.ToLower() == "yes")
                    cc_lsh_difficulty = cc_lsh_difficulty + ", Back";
                if (model.LShSleepIssue?.ToLower() == "yes")
                    cc_lsh_difficulty = cc_lsh_difficulty + ", Sleeping";
                if (!string.IsNullOrEmpty(cc_lsh_difficulty))
                    cc_lsh = cc_lsh + "The patient has difficulty " + cc_lsh_difficulty + " on the left shoulder. ";

                if (model.LShImprove?.Count > 0)
                    cc_lsh = cc_lsh + "There has been improvement with " + string.Join(", ", model.LShImprove) + ".";
                //else
                //    cc_lsh = cc_lsh + "There has been no improvement with physical therapy.";

                //right knee
                if (!string.IsNullOrEmpty(model.RKnPain))
                    cc_rkn = "The patient’s right knee pain level is " + model.RKnPain + "/10. ";
                if (model.RKnSymptoms?.Count > 0)
                    cc_rkn = cc_rkn + "The patient complains of " + string.Join(", ", model.RKnSymptoms) + ". ";


                if (model.RKnReachOverhead?.ToLower() == "yes")
                    cc_rkn_difficulty = "Overhead";
                if (model.RKnReachBack?.ToLower() == "yes")
                    cc_rkn_difficulty = cc_rkn_difficulty + ", Back";
                if (model.RKnSleepIssue?.ToLower() == "yes")
                    cc_rkn_difficulty = cc_rkn_difficulty + ", Sleeping";

                if (!string.IsNullOrEmpty(cc_rkn_difficulty))
                    cc_rkn = cc_rkn + "The patient has difficulty " + cc_rkn_difficulty.TrimStart(',') + " on the right knee. ";

                if (model.RKnImprove?.Count > 0)
                    cc_rkn = cc_rkn + "There has been improvement with " + string.Join(", ", model.RKnImprove) + ".";
                //else
                //    cc_rkn = cc_rkn + "There has been no improvement with physical therapy.";

                //left knee
                if (!string.IsNullOrEmpty(model.LKnPain))
                    cc_lkn = "The patient’s left knee pain level is " + model.LKnPain + "/10. ";
                if (model.LKnSymptoms?.Count > 0)
                    cc_lkn = cc_lkn + "The patient complains of " + string.Join(", ", model.LKnSymptoms) + ". ";


                if (model.LKnReachOverhead?.ToLower() == "yes")
                    cc_lkn_difficulty = "Overhead";
                if (model.LKnReachBack?.ToLower() == "yes")
                    cc_lkn_difficulty = cc_lkn_difficulty + ", Back";
                if (model.LKnSleepIssue?.ToLower() == "yes")
                    cc_lkn_difficulty = cc_lkn_difficulty + ", Sleeping";

                if (!string.IsNullOrEmpty(cc_lkn_difficulty))
                    cc_lkn = cc_lkn + "The patient has difficulty " + cc_lkn_difficulty.TrimStart(',') + " on the left knee. ";

                if (model.LKnImprove?.Count > 0)
                    cc_lkn = cc_lkn + "There has been improvement with " + string.Join(", ", model.LKnImprove) + ".";
                //else
                //    cc_lkn = cc_lkn + "There has been no improvement with physical therapy.";

                return cc_rsh + "<br/>" + cc_lsh + "<br/>" + cc_rkn + "<br/>" + cc_lkn;
            }
        }

        private bool IsPatientPresent(string fname, string lname, DateTime doa)
        {
            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            string cnd = " and fname='" + fname + "' and lname='" + lname + "' and DATE(DOA)='" + doa.ToString("yyyy-MM-dd") + "'  and cmp_id=" + cmpid;


            var data = _patientservices.GetAll(cnd);

            if (data.Count > 0)
                return true;
            else
                return false;
        }

        [HttpGet]
        public IActionResult GeneratePdfBHF(string id, string pdffile = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            Dictionary<string, string> controls = new Dictionary<string, string>();


            ParentService _parentService = new ParentService();


            byte[] pdfBytes = null;
            DataTable dt = _parentService.GetData("select * from vm_patient_ie where intakeid=" + id);
            if (dt.Rows.Count > 0)
            {
                PdfHelper _pdfhelper = new PdfHelper();
                string outputfilename = "";
                string ie_id = "";
                var uploadsFolder = "";
                var filePath = "";
                var signPath = "";
                controls.Add("chk_ie", "Yes");
                try
                {
                    controls.Add("location", dt.Rows[0]["location"].ToString());
                }
                catch { }
                try
                {
                    DataTable dtdos = _parentService.GetData("select id,doe from tbl_patient_ie  where patient_id=" + dt.Rows[0]["patient_id"].ToString());
                    if (dtdos.Rows.Count > 0)
                    {
                        controls.Add("txt_dos", DateTime.Parse(dtdos.Rows[0]["doe"].ToString()).ToString("MM/dd/yyyy"));
                        ie_id = dtdos.Rows[0]["id"].ToString();
                    }
                }
                catch { }


                try
                {
                    DataTable dtbodypart = _parentService.GetData("select bodypart from tbl_ie_page1  where ie_id=" + ie_id);
                    if (dtbodypart.Rows.Count > 0)
                    {
                        string bodypart = dtbodypart.Rows[0]["bodypart"].ToString().ToLower();
                        string[] bodyparts = bodypart.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                                        .Select(x => x.Trim())
                                                        .ToArray();
                        foreach (string data in bodyparts)
                        {
                            controls.Add(data, "Yes");
                        }
                    }
                }
                catch (Exception ex)
                {
                }

                try
                {
                    DataTable dtplan = _parentService.GetData("select FormData from tbl_intake_ai  where id=" + id);
                    if (dtplan.Rows.Count > 0)
                    {
                        string jsonString = dtplan.Rows[0]["FormData"].ToString().ToLower(); ;

                        using JsonDocument doc = JsonDocument.Parse(jsonString);

                        string[] planUTPI = doc.RootElement
                                               .GetProperty("planutpi")
                                               .EnumerateArray()
                                               .Select(x => x.GetString())
                                               .ToArray();
                        foreach (string data in planUTPI)
                        {
                            if (data.Trim() != "")
                                if (!controls.ContainsKey("utpi"))
                                    controls.Add("utpi", "Yes");
                        }
                    }

                }
                catch (Exception ex)
                {

                }


                try
                {
                    uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Downloads/" + cmpclientid);
                    filePath = Path.Combine(uploadsFolder, pdffile);

                    //  signPath = Path.Combine(Directory.GetCurrentDirectory(), "signatures");
                }
                catch (Exception ex)
                {
                    SaveLog(ex, "set Paths");
                }


                try
                {
                    pdfBytes = _pdfhelper.Stamping(filePath, "Id", dt.Rows[0]["patient_id"].ToString(), controls, cmpid, signPath);
                }
                catch (Exception ex)
                {
                    SaveLog(ex, "Pdf Stamping");
                }
                string fileName = $"Superbill.pdf";
                string folder = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments/Others/" + dt.Rows[0]["patient_id"].ToString());

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                string destfilePath = Path.Combine(folder, fileName);

                System.IO.File.WriteAllBytes(destfilePath, pdfBytes);

                //string htmlContent = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "demo.html"));
                // ViewBag.FileName = dt.Rows[0]["LastName"].ToString() + " " + dt.Rows[0]["FirstName"].ToString();


            }

            return File(pdfBytes, "application/pdf", $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Superbill.pdf");
        }


        [HttpGet]
        public IActionResult GeneratePdf(string id, string pdffile = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string cmpclientid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId).ToString();
            Dictionary<string, string> controls = new Dictionary<string, string>();



            ParentService _parentService = new ParentService();


            byte[] pdfBytes = null;
            DataTable dt = _parentService.GetData("select * from vm_patient_ie where intakeid=" + id);
            if (dt.Rows.Count > 0)
            {
                PdfHelper _pdfhelper = new PdfHelper();
                string outputfilename = "";
                string ie_id = "";
                var uploadsFolder = "";
                var filePath = "";
                var signPath = "";
                controls.Add("chk_ie", "Yes");
                try
                {
                    controls.Add("location", dt.Rows[0]["location"].ToString());
                }
                catch { }
                try
                {
                    DataTable dtdos = _parentService.GetData("select id,doe from tbl_patient_ie  where patient_id=" + dt.Rows[0]["patient_id"].ToString());
                    if (dtdos.Rows.Count > 0)
                    {
                        controls.Add("txt_dos", DateTime.Parse(dtdos.Rows[0]["doe"].ToString()).ToString("MM/dd/yyyy"));
                        ie_id = dtdos.Rows[0]["id"].ToString();
                    }
                }
                catch { }


                try
                {
                    DataTable dtbodypart = _parentService.GetData("select bodypart from tbl_ie_page1  where ie_id=" + ie_id);
                    if (dtbodypart.Rows.Count > 0)
                    {
                        string bodypart = dtbodypart.Rows[0]["bodypart"].ToString().ToLower();
                        string[] bodyparts = bodypart.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                                        .Select(x => x.Trim())
                                                        .ToArray();
                        foreach (string data in bodyparts)
                        {
                            controls.Add(data, "Yes");
                        }
                    }
                }
                catch (Exception ex)
                {
                }

                try
                {
                    DataTable dtplan = _parentService.GetData("select FormData from tbl_intake_ai  where id=" + id);
                    if (dtplan.Rows.Count > 0)
                    {
                        string jsonString = dtplan.Rows[0]["FormData"].ToString().ToLower(); ;

                        using JsonDocument doc = JsonDocument.Parse(jsonString);

                        string[] planUTPI = doc.RootElement
                                               .GetProperty("planutpi")
                                               .EnumerateArray()
                                               .Select(x => x.GetString())
                                               .ToArray();
                        foreach (string data in planUTPI)
                        {
                            if (data.Trim() != "")
                                if (!controls.ContainsKey("utpi"))
                                    controls.Add("utpi", "Yes");
                        }
                    }

                }
                catch (Exception ex)
                {

                }


                try
                {
                    uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Downloads/" + cmpclientid);
                    filePath = Path.Combine(uploadsFolder, pdffile);

                    //  signPath = Path.Combine(Directory.GetCurrentDirectory(), "signatures");
                }
                catch (Exception ex)
                {
                    SaveLog(ex, "set Paths");
                }


                try
                {
                    pdfBytes = _pdfhelper.Stamping(filePath, "Id", dt.Rows[0]["patient_id"].ToString(), controls, cmpid, signPath);
                }
                catch (Exception ex)
                {
                    SaveLog(ex, "Pdf Stamping");
                }
                string fileName = $"Superbill.pdf";
                string folder = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments/Others/" + dt.Rows[0]["patient_id"].ToString());

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                string destfilePath = Path.Combine(folder, fileName);

                System.IO.File.WriteAllBytes(destfilePath, pdfBytes);

                //string htmlContent = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "demo.html"));
                // ViewBag.FileName = dt.Rows[0]["LastName"].ToString() + " " + dt.Rows[0]["FirstName"].ToString();


            }

            return File(pdfBytes, "application/pdf", $"{dt.Rows[0]["lname"]}_{dt.Rows[0]["fname"]}_Superbill.pdf");
        }
        #endregion
    }
}
