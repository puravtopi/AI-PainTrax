using AutoMapper;
using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc;
using MS.Models;
using MS.Services;
using Newtonsoft.Json;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using PainTrax.Web.ViewModel;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using static PainTrax.Web.Helper.EnumHelper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using HtmlToOpenXml;
using Microsoft.AspNetCore.Mvc.Rendering;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;
using MailKit;
using System.Net;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class FUVisitController : Controller
    {

        private readonly Common _commonservices = new Common();
        private readonly PatientService _patientservices = new PatientService();
        private readonly PatientFUService _patientFuservices = new PatientFUService();
        private readonly FUPage1Service _fuPage1services = new FUPage1Service();
        private readonly FUPreService _fuPreservices = new FUPreService();
        private readonly FUPostService _fuPostService = new FUPostService();


        private readonly FUPage2Service _fuPage2services = new FUPage2Service();
        private readonly FUPage3Service _fuPage3services = new FUPage3Service();
        private readonly FUNEService _fuNEservices = new FUNEService();
        private readonly FUOtherService _fuOtherService = new FUOtherService();
        private readonly FUCommentService _fuCommentService = new FUCommentService();

        private readonly InscosService _inscosservices = new InscosService();
        private readonly AttorneysService _attorneyservices = new AttorneysService();
        private readonly AadjusterService _aadjusterService = new AadjusterService();
        private readonly EmpService _empService = new EmpService();
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly UserService _userService = new UserService();
        private readonly MacrosMasterService _macroService = new MacrosMasterService();
        private readonly TreatmentMasterService _treatmentService = new TreatmentMasterService();
        private readonly DiagcodesService _diagcodesService = new DiagcodesService();
        private readonly POCServices _pocService = new POCServices();
        private readonly LocationsService _locService = new LocationsService();
        private readonly WebsiteMacrosMasterService _websiteMacrosService = new WebsiteMacrosMasterService();
        private readonly PrintSettingServices _printService = new PrintSettingServices();
        private EnumHelper enumHelper = new EnumHelper();
        private readonly ILogger<VisitController> _logger;
        private IMapper _mapper;
        private readonly ReferringPhysicianService _physicianService = new ReferringPhysicianService();
        private Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;
        string SessionDiffDoc = "true";

        public FUVisitController(ILogger<VisitController> logger, IMapper mapper, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _mapper = mapper;
            _logger = logger;
            Environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Create(int patientIEId, int patientFUId = 0, string type = "")
        {
            VisitVM obj = new VisitVM();

            try
            {
                int id = patientIEId;
                obj.vaccinated = false;
                obj.mc = "";

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                ViewBag.locList = _commonservices.GetLocations(cmpid.Value);

                var macroList = _websiteMacrosService.GetAutoComplete(cmpid.Value);
                ViewBag.csList = _commonservices.GetCaseType(cmpid.Value);
                ViewBag.asList = _commonservices.GetAccidenttype(cmpid.Value);

                ViewBag.macroList = JsonConvert.SerializeObject(macroList);
                ViewBag.physicianList = _commonservices.GetPhysician(cmpid.Value);
                ViewBag.stateList = _commonservices.GetState(cmpid.Value);


                macroList = _websiteMacrosService.GetAutoComplete(cmpid.Value, "CC");
                ViewBag.ccmacroList = JsonConvert.SerializeObject(macroList);

                macroList = _websiteMacrosService.GetAutoComplete(cmpid.Value, "PE");
                ViewBag.pemacroList = JsonConvert.SerializeObject(macroList);

                var providers = _userService.GetProviders(cmpid.Value);
                ViewBag.providerList = providers;

                ViewBag.insuranceList = _inscosservices.GetAautoComplete("");

                ViewBag.type = type;

                if (patientFUId > 0)
                {
                    var ieData = _ieService.GetOnebyPatientId(patientIEId);

                    int patientId = 0;

                    if (ieData != null)
                    {
                        obj.id = ieData.id;
                        patientId = ieData.patient_id.Value;
                        obj.locationid = ieData.location_id;
                        obj.providerid = ieData.provider_id;
                        obj.location = ieData.location;
                        obj.doa = ieData.doa;
                        obj.doi = ieData.doe;
                        obj.old_dos = ieData.doe;
                        obj.prime_claim_no = ieData.primary_claim_no;
                        obj.prime_policy_no = ieData.primary_policy_no;
                        obj.prime_WCB_group = ieData.primary_wcb_group;

                        obj.sec_claim_no = ieData.secondary_claim_no;
                        obj.sec_policy_no = ieData.secondary_policy_no;
                        obj.sec_WCB_group = ieData.secondary_wcb_group;
                        obj.alert_note = ieData.alert_note;
                        //obj.referring_physician = ieData.referring_physician;
                        obj.physicianid = ieData.physicianid;
                        obj.casetype = ieData.casetype;
                        obj.compensation = ieData.compensation;
                        obj.accidentType = ieData.accidentType;
                    }

                    var fuData = _patientFuservices.GetOne(patientFUId);

                    obj.dos = fuData.doe;
                    obj.dov = fuData.doe;
                    obj.accidentType = fuData.accident_type;
                    obj.physicianid = fuData.physicianid;
                    obj.providerid = fuData.provider_id;
                    obj.locationid = fuData.location_id == null ? ieData.location_id : fuData.location_id;
                    obj.type = fuData.type;
                    obj.procedure_performed = fuData.procedure_performed;
                    if (ieData.primary_ins_cmp_id != null)
                    {
                        var primaryCmp = _inscosservices.GetOne(ieData.primary_ins_cmp_id.Value);

                        if (primaryCmp != null)
                        {
                            obj.prime_cmpname = primaryCmp.cmpname;
                            obj.prime_address = primaryCmp.address1;
                            obj.prime_phone = primaryCmp.telephone;
                        }
                    }

                    if (ieData.secondary_ins_cmp_id != null)
                    {
                        var secondaryCmp = _inscosservices.GetOne(ieData.secondary_ins_cmp_id.Value);

                        if (secondaryCmp != null)
                        {
                            obj.sec_cmpname = secondaryCmp.cmpname;
                            obj.sec_address = secondaryCmp.address1;
                            obj.sec_phone = secondaryCmp.telephone;
                        }
                    }

                    if (ieData.emp_id != null)
                    {
                        var empData = _empService.GetOne(ieData.emp_id.Value);

                        if (empData != null)
                        {
                            obj.emp_address = empData.address;
                            obj.emp_fax_no = empData.fax;
                            obj.emp_name = empData.name;
                            obj.emp_phone = empData.phone;
                        }
                    }

                    if (ieData.adjuster_id != null)
                    {
                        var adjData = _aadjusterService.GetOne(ieData.adjuster_id.Value);

                        if (adjData != null)
                        {
                            obj.adj_email = adjData.emailaddress;
                            obj.adj_fax_no = adjData.fax;
                            obj.adj_name = adjData.adjuster;
                            obj.adj_phone = adjData.telephone;
                        }
                    }

                    if (ieData.attorney_id != null)
                    {
                        var attrData = _attorneyservices.GetOne(ieData.attorney_id.Value);

                        if (attrData != null)
                        {
                            obj.attory_email = attrData.EmailId;
                            obj.attory_fax_no = "";
                            obj.attory_name = attrData.Attorney;
                            obj.attory_phone = attrData.ContactNo;
                            obj.paralegal = attrData.Paralegal;
                        }
                    }

                    var page1Data = _fuPage1services.GetOne(patientFUId);

                    if (page1Data != null)
                    {
                        var page1IEData = _mapper.Map<tbl_ie_page1>(page1Data);
                        obj.Page1 = page1IEData;

                        string daignoLink = "";
                        if (!string.IsNullOrEmpty(page1Data.bodypart))
                        {
                            for (int i = 0; i < page1Data.bodypart.Split(',').Length; i++)
                            {
                                var linkbody = page1Data.bodypart.Split(',')[i].Replace(" ", "_");
                                daignoLink += "<a href='javascript:void(0)' onclick=openDaignoModel('" + linkbody + "')>" + page1Data.bodypart.Split(',')[i] + "</a><br/>";
                            }

                        }
                        ViewBag.DaignoLink = daignoLink;
                    }
                    else
                        obj.Page1 = new tbl_ie_page1();

                    // for pre op start
                    var preData = _fuPreservices.GetOne(patientFUId);
                    if (preData != null)
                    {
                        var preData1 = _mapper.Map<tbl_pre>(preData);
                        obj.pre = preData1;
                    }
                    else
                    {



                        tbl_pre objdefaultdata = new tbl_pre();
                        var defData = _fuPreservices.GetOneOPDefault(patientIEId);
                        if (defData != null)
                        {
                            objdefaultdata.txtPastMedicalHistory = defData.txtPastMedicalHistory;
                            objdefaultdata.txtpastsurgicalhistory = defData.txtpastsurgicalhistory;
                            objdefaultdata.txtdailyMedications = defData.txtdailyMedications;
                            objdefaultdata.txtAllergies = defData.txtAllergies;
                            objdefaultdata.txtSH = defData.txtSH;
                            objdefaultdata.txtFamilyHistory = defData.txtFamilyHistory;
                            objdefaultdata.txtPresentillness = defData.txtPresentillness;
                        }

                        var preDatadef = _mapper.Map<tbl_pre>(objdefaultdata);
                        obj.pre = preDatadef;

                    }

                    // for post op start
                    var postData = _fuPostService.GetOne(patientFUId);
                    if (postData != null)
                    {
                        var postData1 = _mapper.Map<tbl_post>(postData);
                        obj.post = postData1;
                    }
                    else
                        obj.post = new tbl_post();


                    var page2Data = _fuPage2services.GetOne(patientFUId);

                    if (page2Data != null)
                    {
                        var page2IEData = _mapper.Map<tbl_ie_page2>(page2Data);
                        obj.Page2 = page2IEData;
                    }
                    else
                        obj.Page2 = new tbl_ie_page2();

                    var page3Data = _fuPage3services.GetOne(patientFUId);

                    if (page3Data != null)
                    {
                        var page3IEData = _mapper.Map<tbl_ie_page3>(page3Data);
                        obj.Page3 = page3IEData;
                    }
                    else
                        obj.Page3 = new tbl_ie_page3();


                    var NEData = _fuNEservices.GetOne(patientFUId);

                    if (NEData != null)
                    {
                        var neIEData = _mapper.Map<tbl_ie_ne>(NEData);
                        obj.NE = neIEData;
                    }
                    else
                        obj.NE = new tbl_ie_ne();


                    var pageOther = _fuOtherService.GetOne(patientFUId);

                    if (pageOther != null)
                    {
                        var otherIEData = _mapper.Map<tbl_ie_other>(pageOther);

                        obj.Other = otherIEData;

                        var tretmentdesc = string.IsNullOrEmpty(pageOther.treatment_delimit_desc) ? null : pageOther.treatment_delimit_desc.Split('^');


                        var _data = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);

                        List<tbl_treatment_master> lst = new List<tbl_treatment_master>();
                        int i = 0;
                        foreach (var item in _data)
                        {
                            if (!string.IsNullOrEmpty(pageOther.treatment_delimit))
                            {
                                if (pageOther.treatment_delimit.Contains(item.id.ToString()))
                                {
                                    if (tretmentdesc != null)
                                    {
                                        if (tretmentdesc.Length >= (i + 1))
                                            item.treatment_details = tretmentdesc[i];
                                    }
                                    item.pre_select = true;
                                    i++;
                                }
                                else
                                    item.pre_select = false;
                            }

                            lst.Add(item);

                        }

                        obj.Other.listTreatmentMaster = lst;
                    }
                    else
                    {
                        obj.Other = new tbl_ie_other();
                        var _data = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);
                        obj.Other.followup_duration = "2 weeks.";
                        obj.Other.listTreatmentMaster = _data;
                    }

                    var pageComment = _ieService.GetOneComment(patientIEId);

                    if (pageComment != null)
                    {
                        obj.Comment = pageComment;
                    }
                    else
                        obj.Comment = new tbl_ie_comment();


                    var patientData = _patientservices.GetOne(patientId);

                    if (patientData != null)
                    {
                        obj.fname = patientData.fname;
                        obj.lname = patientData.lname;
                        obj.mname = patientData.mname;
                        obj.gender = patientData.gender;
                        obj.dob = patientData.dob;
                        obj.email = patientData.email;
                        obj.handeness = patientData.handeness;
                        obj.ssn = patientData.ssn;
                        obj.address = patientData.address;
                        obj.city = patientData.city;
                        obj.state = patientData.state;
                        obj.zip = patientData.zip;
                        obj.account_no = patientData.account_no;
                        obj.home_ph = patientData.home_ph;
                        obj.mobile = patientData.mobile;
                        obj.mc = patientData.mc;
                        obj.vaccinated = patientData.vaccinated;
                        obj.patientid = patientData.id;
                        obj.age = patientData.age;
                        obj.physicianid = patientData.physicianid;
                    }
                    //for documents
                    {

                        HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, ieData.patient_id.Value);

                        var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments");
                        bool folderExists = Directory.Exists(FolderPath);
                        if (!folderExists)
                            Directory.CreateDirectory(FolderPath);

                        string[] dirs = Directory.GetDirectories(FolderPath, "*", SearchOption.TopDirectoryOnly);

                        List<TreeViewNode> nodes = new List<TreeViewNode>();

                        int i = 0;
                        foreach (var item in dirs)
                        {
                            i++;
                            string FolderName = System.IO.Path.GetFileName(item);
                            var FolderPathFile = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", FolderName.ToString(), ieData.patient_id.Value.ToString());
                            int j = 0;

                            bool folderExistsNew = Directory.Exists(FolderPathFile);
                            if (!folderExistsNew)
                                Directory.CreateDirectory(FolderPathFile);

                            foreach (var item1 in Directory.GetFiles(FolderPathFile))
                            {
                                j++;
                                nodes.Add(new TreeViewNode { id = j.ToString() + "-" + i.ToString() + "~" + FolderName.ToString() + "$" + System.IO.Path.GetFileName(item1), parent = i.ToString(), text = System.IO.Path.GetFileName(item1) });
                            }

                            nodes.Add(new TreeViewNode { id = i.ToString(), parent = "#", text = FolderName.ToString() + "(" + j + ")" });
                        }

                        obj.doc_json = JsonConvert.SerializeObject(nodes, Formatting.Indented);

                    }
                }
                else if (id > 0)
                {

                    var ieData = _ieService.GetOnebyPatientId(id);

                    if (ieData != null)
                    {
                        obj.id = ieData.id;

                        obj.locationid = ieData.location_id;
                        obj.location = ieData.location;
                        obj.providerid = ieData.provider_id;
                        obj.doa = ieData.doa;
                        obj.dos = ieData.doe;
                        obj.prime_claim_no = ieData.primary_claim_no;
                        obj.prime_policy_no = ieData.primary_policy_no;
                        obj.prime_WCB_group = ieData.primary_wcb_group;

                        obj.sec_claim_no = ieData.secondary_claim_no;
                        obj.sec_policy_no = ieData.secondary_policy_no;
                        obj.sec_WCB_group = ieData.secondary_wcb_group;
                        obj.alert_note = ieData.alert_note;
                        //obj.referring_physician = ieData.referring_physician;
                        obj.physicianid = ieData.physicianid;
                        obj.compensation = ieData.compensation;
                        obj.casetype = ieData.casetype;
                    }
                    //if (obj.compensation != null)
                    //    obj.casetype = System.Enum.GetName(typeof(CaseType), Convert.ToInt32(obj.compensation));

                    if (ieData.primary_ins_cmp_id != null)
                    {
                        var primaryCmp = _inscosservices.GetOne(ieData.primary_ins_cmp_id.Value);

                        if (primaryCmp != null)
                        {
                            obj.prime_cmpname = primaryCmp.cmpname;
                            obj.prime_address = primaryCmp.address1;
                            obj.prime_phone = primaryCmp.telephone;
                        }
                    }

                    if (ieData.secondary_ins_cmp_id != null)
                    {
                        var secondaryCmp = _inscosservices.GetOne(ieData.secondary_ins_cmp_id.Value);

                        if (secondaryCmp != null)
                        {
                            obj.sec_cmpname = secondaryCmp.cmpname;
                            obj.sec_address = secondaryCmp.address1;
                            obj.sec_phone = secondaryCmp.telephone;
                        }
                    }

                    if (ieData.emp_id != null)
                    {
                        var empData = _empService.GetOne(ieData.emp_id.Value);

                        if (empData != null)
                        {
                            obj.emp_address = empData.address;
                            obj.emp_fax_no = empData.fax;
                            obj.emp_name = empData.name;
                            obj.emp_phone = empData.phone;
                        }
                    }

                    if (ieData.adjuster_id != null)
                    {
                        var adjData = _aadjusterService.GetOne(ieData.adjuster_id.Value);

                        if (adjData != null)
                        {
                            obj.adj_email = adjData.emailaddress;
                            obj.adj_fax_no = adjData.fax;
                            obj.adj_name = adjData.adjuster;
                            obj.adj_phone = adjData.telephone;
                        }
                    }

                    if (ieData.attorney_id != null)
                    {
                        var attrData = _attorneyservices.GetOne(ieData.attorney_id.Value);

                        if (attrData != null)
                        {
                            obj.attory_email = attrData.EmailId;
                            obj.attory_fax_no = "";
                            obj.attory_name = attrData.Attorney;
                            obj.attory_phone = attrData.ContactNo;
                        }
                    }

                    var page1Data = _ieService.GetOnePage1(id);

                    if (page1Data != null)
                    {
                        obj.Page1 = page1Data;

                        string daignoLink = "";
                        if (!string.IsNullOrEmpty(page1Data.bodypart))
                        {
                            for (int i = 0; i < page1Data.bodypart.Split(',').Length; i++)
                            {
                                var linkbody = page1Data.bodypart.Split(',')[i].Replace(" ", "_");
                                daignoLink += "<a href=# onclick=openDaignoModel('" + linkbody + "')>" + page1Data.bodypart.Split(',')[i] + "</a><br/>";
                            }

                        }
                        ViewBag.DaignoLink = daignoLink;
                    }
                    else
                        obj.Page1 = new tbl_ie_page1();




                    // preop  start. 

                    tbl_pre objdefaultdata = new tbl_pre();
                    var defData = _fuPreservices.GetOneOPDefault(patientIEId);
                    if (defData != null)
                    {
                        objdefaultdata.txtPastMedicalHistory = defData.txtPastMedicalHistory;
                        objdefaultdata.txtpastsurgicalhistory = defData.txtpastsurgicalhistory;
                        objdefaultdata.txtdailyMedications = defData.txtdailyMedications;
                        objdefaultdata.txtAllergies = defData.txtAllergies;
                        objdefaultdata.txtFamilyHistory = defData.txtFamilyHistory;
                        objdefaultdata.txtSH = defData.txtSH;
                        objdefaultdata.txtPresentillness = defData.txtPresentillness;
                    }

                    var preDatadef = _mapper.Map<tbl_pre>(objdefaultdata);
                    obj.pre = preDatadef;

                    // preop end.
                    var postData = _fuPostService.GetOne(patientFUId);
                    if (postData != null)
                    {
                        var postData1 = _mapper.Map<tbl_post>(postData);
                        obj.post = postData1;
                    }
                    else
                        obj.post = new tbl_post();




                    var page2Data = _ieService.GetOnePage2(id);

                    if (page2Data != null)
                    {
                        obj.Page2 = page2Data;
                    }
                    else
                        obj.Page2 = new tbl_ie_page2();

                    var NEData = _ieService.GetOneNE(id);

                    if (NEData != null)
                    {
                        obj.NE = NEData;
                    }
                    else
                        obj.NE = new tbl_ie_ne();

                    var page3Data = _ieService.GetOnePage3(id);

                    if (page3Data != null)
                    {
                        obj.Page3 = page3Data;
                    }
                    else
                        obj.Page3 = new tbl_ie_page3();


                    var pageComment = _ieService.GetOneComment(id);

                    if (pageComment != null)
                    {
                        obj.Comment = pageComment;
                    }
                    else
                        obj.Comment = new tbl_ie_comment();

                    var pageOther = _ieService.GetOneOtherPage(id);

                    if (pageOther != null)
                    {
                        obj.Other = pageOther;
                        var _data = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);

                        List<tbl_treatment_master> lst = new List<tbl_treatment_master>();
                        foreach (var item in _data)
                        {
                            if (!string.IsNullOrEmpty(pageOther.treatment_delimit))
                            {
                                if (pageOther.treatment_delimit.Contains(item.id.ToString()))
                                    item.pre_select = true;
                                else
                                    item.pre_select = false;
                            }

                            lst.Add(item);
                        }

                        obj.Other.listTreatmentMaster = lst;
                    }
                    else
                    {
                        obj.Other = new tbl_ie_other();
                        var _data = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);
                        obj.Other.followup_duration = "2 weeks.";
                        obj.Other.listTreatmentMaster = _data;
                    }
                    var patientData = _patientservices.GetOne(ieData.patient_id.Value);

                    if (patientData != null)
                    {
                        obj.fname = patientData.fname;
                        obj.lname = patientData.lname;
                        obj.mname = patientData.mname;
                        obj.gender = patientData.gender;
                        obj.dob = patientData.dob;
                        obj.email = patientData.email;
                        obj.handeness = patientData.handeness;
                        obj.ssn = patientData.ssn;
                        obj.address = patientData.address;
                        obj.city = patientData.city;
                        obj.state = patientData.state;
                        obj.zip = patientData.zip;
                        obj.account_no = patientData.account_no;
                        obj.home_ph = patientData.home_ph;
                        obj.mobile = patientData.mobile;
                        obj.mc = patientData.mc;
                        obj.vaccinated = patientData.vaccinated;
                        obj.patientid = patientData.id;
                        obj.physicianid = patientData.physicianid;
                    }
                    //for documents
                    {

                        HttpContext.Session.SetInt32(SessionKeys.SessionPatientId, ieData.patient_id.Value);

                        var FolderPath = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments");
                        bool folderExists = Directory.Exists(FolderPath);
                        if (!folderExists)
                            Directory.CreateDirectory(FolderPath);

                        string[] dirs = Directory.GetDirectories(FolderPath, "*", SearchOption.TopDirectoryOnly);

                        List<TreeViewNode> nodes = new List<TreeViewNode>();

                        int i = 0;
                        foreach (var item in dirs)
                        {
                            i++;
                            string FolderName = System.IO.Path.GetFileName(item);
                            var FolderPathFile = Path.Combine(Directory.GetCurrentDirectory(), "PatientDocuments", FolderName.ToString(), ieData.patient_id.Value.ToString());
                            int j = 0;

                            bool folderExistsNew = Directory.Exists(FolderPathFile);
                            if (!folderExistsNew)
                                Directory.CreateDirectory(FolderPathFile);

                            foreach (var item1 in Directory.GetFiles(FolderPathFile))
                            {
                                j++;
                                nodes.Add(new TreeViewNode { id = j.ToString() + "-" + i.ToString() + "~" + FolderName.ToString() + "$" + System.IO.Path.GetFileName(item1), parent = i.ToString(), text = System.IO.Path.GetFileName(item1) });
                            }

                            nodes.Add(new TreeViewNode { id = i.ToString(), parent = "#", text = FolderName.ToString() + "(" + j + ")" });
                        }

                        obj.doc_json = JsonConvert.SerializeObject(nodes, Formatting.Indented);

                    }
                }
                else
                {
                    obj.patientid = 0;
                    obj.id = 0;
                    obj.Page1 = new tbl_ie_page1();
                    obj.Page2 = new tbl_ie_page2();
                    obj.post = new tbl_post();
                    obj.Page3 = new tbl_ie_page3();
                    obj.NE = new tbl_ie_ne();
                    obj.Comment = new tbl_ie_comment();
                    obj.Other = new tbl_ie_other();
                    obj.Other.followup_duration = "2-4 weeks.";
                    obj.Page3.gait = "Guarded";

                    var _data = _treatmentService.GetAll(" and cmp_id=" + cmpid.Value);

                    obj.Other.listTreatmentMaster = _data;

                }

                obj.listWebsiteMacros = macroList;

                obj.fu_id = patientFUId;

                if (obj.Page3 != null)
                {
                    obj.Page3.diagcervialbulge_study = (obj.Page3.diagcervialbulge_study == null) ? "1" : obj.Page3.diagcervialbulge_study;
                    obj.Page3.diagthoracicbulge_study = (obj.Page3.diagthoracicbulge_study == null) ? "1" : obj.Page3.diagthoracicbulge_study;
                    obj.Page3.diaglumberbulge_study = (obj.Page3.diaglumberbulge_study == null) ? "1" : obj.Page3.diaglumberbulge_study;
                    obj.Page3.diagleftshoulder_study = (obj.Page3.diagleftshoulder_study == null) ? "1" : obj.Page3.diagleftshoulder_study;
                    obj.Page3.diagrightshoulder_study = (obj.Page3.diagrightshoulder_study == null) ? "1" : obj.Page3.diagrightshoulder_study;
                    obj.Page3.diagleftknee_study = (obj.Page3.diagleftknee_study == null) ? "1" : obj.Page3.diagleftknee_study;
                    obj.Page3.diagrightknee_study = (obj.Page3.diagrightknee_study == null) ? "1" : obj.Page3.diagrightknee_study;
                    obj.Page3.diaglumberbulge_study = (obj.Page3.diaglumberbulge_study == null) ? "1" : obj.Page3.diaglumberbulge_study;

                    obj.Page3.other1_study = (obj.Page3.other1_study == null) ? "0" : obj.Page3.other1_study;
                    obj.Page3.other2_study = (obj.Page3.other2_study == null) ? "0" : obj.Page3.other2_study;
                    obj.Page3.other3_study = (obj.Page3.other3_study == null) ? "0" : obj.Page3.other3_study;
                    obj.Page3.other4_study = (obj.Page3.other4_study == null) ? "0" : obj.Page3.other4_study;
                    obj.Page3.other5_study = (obj.Page3.other5_study == null) ? "0" : obj.Page3.other5_study;
                    obj.Page3.other6_study = (obj.Page3.other6_study == null) ? "0" : obj.Page3.other6_study;
                    obj.Page3.other7_study = (obj.Page3.other7_study == null) ? "0" : obj.Page3.other7_study;
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Create");
            }
            return View(obj);
        }

        [HttpPost]
        public IActionResult SaveDetails(VisitVM model)
        {
            try
            {
                var data = model;
                var _isExist = _patientFuservices.IsFuExist(model.fu_id.Value, model.dov == null ? null : model.dov.Value.ToString("yyyy-MM-dd"), model.type, model.id.Value);

                if (!_isExist)
                {
                    int patientId = 0, priminsId = 0, secinsId = 0, attornyId = 0, adjusterId = 0, empId = 0;

                    int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    int? userid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId);
                    int age = calculateAge(model.dob.Value, model.dov);

                    tbl_patient objPatient = new tbl_patient()
                    {
                        account_no = model.account_no,
                        address = model.address,
                        city = model.city,
                        dob = model.dob,
                        email = model.email,
                        fname = model.fname,
                        gender = model.gender,
                        home_ph = model.home_ph,
                        lname = model.lname,
                        mc = model.mc,
                        mname = model.mname,
                        mobile = model.mobile,
                        ssn = model.ssn,
                        state = model.state,
                        physicianid = model.physicianid,
                        vaccinated = model.vaccinated,
                        zip = model.zip,
                        cmp_id = cmpid,
                        createdby = userid,
                        age = age
                    };

                    if (model.id.Value > 0)
                    {
                        objPatient.id = model.patientid;
                        _patientservices.Update(objPatient);
                        patientId = model.patientid.Value;
                    }
                    else
                    {
                        patientId = _patientservices.Insert(objPatient);
                    }
                    var query = "";
                    List<tbl_inscos> insdata = new List<tbl_inscos>();
                    tbl_inscos objInscos = new tbl_inscos();

                    if (!string.IsNullOrEmpty(model.prime_cmpname))
                    {
                        query = " and cmpname='" + model.prime_cmpname + "' and cmp_id=" + cmpid + "";

                        insdata = _inscosservices.GetAll(query);

                        //save primary insurance

                        objInscos = new tbl_inscos()
                        {
                            address1 = model.prime_address,
                            cmpname = model.prime_cmpname,
                            telephone = model.prime_phone,
                            cmp_id = cmpid,
                            createdby = userid

                        };

                        if (insdata.Count > 0)
                        {
                            objInscos.id = insdata[0].id.Value;
                            _inscosservices.Update(objInscos);
                            priminsId = insdata[0].id.Value;
                        }
                        else
                        {
                            priminsId = _inscosservices.Insert(objInscos);
                        }

                    }

                    if (!string.IsNullOrEmpty(model.sec_cmpname))
                    {

                        query = " and cmpname='" + model.sec_cmpname + "' and cmp_id=" + cmpid + "";

                        insdata = _inscosservices.GetAll(query);

                        //save secondary insurance

                        objInscos = new tbl_inscos()
                        {
                            address1 = model.sec_address,
                            cmpname = model.sec_cmpname,
                            telephone = model.sec_phone,
                            cmp_id = cmpid,
                            createdby = userid

                        };

                        if (insdata.Count > 0)
                        {
                            objInscos.id = insdata[0].id.Value;
                            _inscosservices.Update(objInscos);
                            secinsId = insdata[0].id.Value;
                        }
                        else
                        {
                            secinsId = _inscosservices.Insert(objInscos);
                        }
                    }

                    if (!string.IsNullOrEmpty(model.attory_name))
                    {

                        query = " and Attorney='" + model.attory_name + "' and cmp_id=" + cmpid + "";

                        var attrydata = _attorneyservices.GetAll(query);


                        //save attorney

                        tbl_attorneys objAttorneys = new tbl_attorneys()
                        {
                            Attorney = model.attory_name,
                            EmailId = model.attory_email,
                            ContactNo = model.attory_phone,
                            cmp_id = cmpid,
                            CreatedBy = userid,
                            Paralegal = model.paralegal

                        };

                        if (attrydata.Count > 0)
                        {
                            objAttorneys.Id = attrydata[0].Id.Value;
                            _attorneyservices.Update(objAttorneys);
                            attornyId = attrydata[0].Id.Value;
                        }
                        else
                        {
                            attornyId = _attorneyservices.Insert(objAttorneys);
                        }
                    }

                    if (!string.IsNullOrEmpty(model.adj_name))
                    {
                        query = " and adjuster='" + model.adj_name + "' and cmp_id=" + cmpid + "";

                        var adjdata = _aadjusterService.GetAll(query);

                        //save adjuster

                        tbl_adjuster objAdjuster = new tbl_adjuster()
                        {
                            adjuster = model.adj_name,
                            emailaddress = model.adj_email,
                            telephone = model.adj_phone,
                            fax = model.adj_fax_no,
                            cmp_id = cmpid,
                            created_by = userid
                        };

                        if (adjdata.Count > 0)
                        {
                            objAdjuster.id = adjdata[0].id.Value;
                            _aadjusterService.Update(objAdjuster);
                            adjusterId = adjdata[0].id.Value;
                        }
                        else
                        {
                            adjusterId = _aadjusterService.Insert(objAdjuster);
                        }
                    }


                    if (!string.IsNullOrEmpty(model.emp_name))
                    {
                        query = " and name='" + model.emp_name + "' and patient_id=" + model.patientid + "";

                        var empdata = _empService.GetAll(query);

                        //save employee

                        tbl_emp objEmp = new tbl_emp()
                        {
                            address = model.emp_address,
                            name = model.emp_name,
                            phone = model.emp_phone,
                            fax = model.emp_fax_no,
                            patient_id = patientId
                        };

                        if (empdata.Count > 0)
                        {
                            objEmp.id = empdata[0].id.Value;
                            _empService.Update(objEmp);
                            empId = empdata[0].id.Value;
                        }
                        else
                        {
                            empId = _empService.Insert(objEmp);
                        }
                    }

                    //save IE details 

                    tbl_patient_ie objIE = new tbl_patient_ie()
                    {
                        adjuster_id = adjusterId,
                        attorney_id = attornyId,
                        created_by = userid,
                        doa = model.doa,
                        emp_id = empId,
                        is_active = true,

                        provider_id = model.providerid,
                        patient_id = patientId,
                        primary_claim_no = model.prime_claim_no,
                        primary_ins_cmp_id = priminsId,
                        primary_policy_no = model.prime_policy_no,
                        primary_wcb_group = model.prime_WCB_group,
                        secondary_claim_no = model.sec_claim_no,
                        secondary_ins_cmp_id = secinsId,
                        secondary_policy_no = model.sec_policy_no,
                        secondary_wcb_group = model.sec_WCB_group,
                        compensation = model.compensation,
                        alert_note = model.alert_note,
                        referring_physician = model.referring_physician

                    };
                    int ie = 0;
                    if (model.id.Value > 0)
                    {
                        objIE.id = model.id.Value;
                        _ieService.UpdateFromFU(objIE);
                        ie = model.id.Value;
                    }
                    else
                        ie = _ieService.Insert(objIE);


                    //save FU details 

                    tbl_patient_fu objFU = new tbl_patient_fu()
                    {

                        created_by = userid,
                        doe = model.dov,
                        patientIE_ID = ie,
                        cmp_id = cmpid,
                        created_date = System.DateTime.Now,
                        is_active = true,
                        patient_id = patientId,
                        extra_comments = model.alert_note,
                        type = model.type,
                        accident_type = model.accidentType,
                        provider_id = model.providerid,
                        physicianid = model.physicianid,
                        location_id = model.locationid
                    };
                    int fu_id = 0;
                    if (model.fu_id.Value > 0)
                    {
                        objFU.id = model.fu_id.Value;
                        _patientFuservices.Update(objFU);
                        fu_id = model.fu_id.Value;

                        if (model.old_dos != null)
                        {
                            _patientFuservices.UpdateProcedureExecuteDate(model.old_dos.Value.ToString("yyyy-MM-dd"),
                          model.dov.Value.ToString("yyyy-MM-dd"), fu_id.ToString());
                        }
                    }
                    else
                        fu_id = _patientFuservices.Insert(objFU);

                    HttpContext.Session.SetInt32(SessionKeys.SessionIEId, ie);
                    return Json(new { status = 1, patintid = patientId, ieid = ie, fuid = fu_id });
                }
                else
                {
                    return Json(new { status = -1 });
                }
            }

            catch (Exception ex)
            {
                SaveLog(ex, "SaveDetails");

            }
            return Json(new { status = 0 });
        }

        [HttpPost]
        public IActionResult SavePage1(tbl_fu_page1 model)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;

                var obj = _fuPage1services.GetOne(model.fu_id.Value);

                if (obj != null)
                {
                    model.id = obj.id;
                    _fuPage1services.Update(model);
                }
                else
                    _fuPage1services.Insert(model);

            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePage1");
            }
            return Json(1);
        }

        [HttpPost]
        public IActionResult SavePre(tbl_pre model)
        {
            try
            {
                //int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                //model.cmp_id = cmpid;

                var obj = _fuPreservices.GetOne(model.PatientFU_ID.Value);
                if (Convert.ToBoolean(model.chkLeftShoulder.Value) || Convert.ToBoolean(model.chkRightShoulder.Value) || Convert.ToBoolean(model.chkLeftKnee.Value) || Convert.ToBoolean(model.chkRightKnee.Value))
                {
                    if (obj != null)
                    {
                        model.id = obj.id;
                        _fuPreservices.Update(model);
                    }
                    else
                        _fuPreservices.Insert(model);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePage1");
            }
            return Json(1);
        }
        [HttpPost]
        public IActionResult SavePost(tbl_post model)
        {
            try
            {
                //int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                //model.cmp_id = cmpid;

                var obj = _fuPostService.GetOne(model.PatientFU_ID.Value);
                if (Convert.ToBoolean(model.chkLeftShoulder.Value) || Convert.ToBoolean(model.chkRightShoulder.Value) || Convert.ToBoolean(model.chkLeftKnee.Value) || Convert.ToBoolean(model.chkRightKnee.Value))
                {
                    if (obj != null)
                    {
                        model.id = obj.id;
                        _fuPostService.Update(model);
                    }
                    else
                        _fuPostService.Insert(model);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePost");
            }
            return Json(1);
        }


        [HttpPost]
        public IActionResult SavePage2(tbl_fu_page2 model)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;

                var obj = _fuPage2services.GetOne(model.fu_id.Value);

                if (obj != null)
                {
                    model.id = obj.id;
                    _fuPage2services.Update(model);
                }
                else
                    _fuPage2services.Insert(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePage2");

            }
            return Json(1);
        }

        [HttpPost]
        public IActionResult SavePage3(tbl_fu_page3 model, int ieid, int fuid, int patientid, int? id, string discharge_medications)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;

                model.ie_id = ieid;
                model.patient_id = patientid;
                model.fu_id = fuid;

                if (model.diagcervialbulge_date != null)
                    model.diagcervialbulge = true;
                else
                    model.diagcervialbulge = false;

                if (model.diagleftknee_date != null)
                    model.diagleftknee = true;
                else
                    model.diagleftknee = false;

                if (model.diaglumberbulge_date != null)
                    model.diaglumberbulge = true;
                else
                    model.diaglumberbulge = false;

                if (model.diagrightknee_date != null)
                    model.diagrightknee = true;
                else
                    model.diagrightknee = false;

                if (model.diagthoracicbulge_date != null)
                    model.diagthoracicbulge = true;
                else
                    model.diagthoracicbulge = false;

                if (model.diagrightshoulder_date != null)
                    model.diagrightshoulder = true;
                else
                    model.diagrightshoulder = false;

                if (model.diagleftshoulder_date != null)
                    model.diagleftshoulder = true;
                else
                    model.diagleftshoulder = false;

                if (model.other1_date != null)
                    model.other1 = true;
                else
                    model.other1 = false;

                if (model.other2_date != null)
                    model.other2 = true;
                else
                    model.other2 = false;

                if (model.other3_date != null)
                    model.other3 = true;
                else
                    model.other3 = false;

                if (model.other4_date != null)
                    model.other4 = true;
                else
                    model.other4 = false;

                if (model.other5_date != null)
                    model.other5 = true;
                else
                    model.other5 = false;

                if (model.other6_date != null)
                    model.other6 = true;
                else
                    model.other6 = false;

                if (model.other7_date != null)
                    model.other7 = true;
                else
                    model.other7 = false;
                int data = 1;


                var obj = _fuPage3services.GetOne(model.fu_id.Value);

                if (obj != null)
                {
                    model.id = obj.id;
                    _fuPage3services.Update(model);
                }
                else
                    _fuPage3services.Insert(model);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePage3");
            }
            return Json(1);

        }

        [HttpPost]
        public IActionResult SaveNE(tbl_fu_ne model)
        {
            try
            {
                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                model.cmp_id = cmpid;
                var data = 1;

                var obj = _fuNEservices.GetOne(model.fu_id.Value);

                if (obj != null)
                {
                    model.id = obj.id;
                    _fuNEservices.Update(model);
                }
                else
                    _fuNEservices.Insert(model);


            }
            catch (Exception ex)
            {
                SaveLog(ex, "SaveNE");
            }
            return Json(1);
        }

        [HttpPost]
        public IActionResult SaveOther(tbl_fu_other model)
        {
            try
            {
                var obj = _fuOtherService.GetOne(model.fu_id.Value);

                if (obj != null)
                {
                    model.id = obj.id;
                    _fuOtherService.Update(model);
                }
                else
                    _fuOtherService.Insert(model);


            }
            catch (Exception ex)
            {
                SaveLog(ex, "SaveOther");

            }
            return Json(1);
        }

        [HttpPost]
        public IActionResult SaveComment(tbl_ie_comment model)
        {

            var data = 1;
            if (model.id > 0)
            {
                data = model.id.Value;
                _ieService.UpdateComment(model);
            }
            else
                _ieService.InsertComment(model);

            return Json(data);
        }
        public IActionResult LoadCommentPartial(int ie_id)
        {
            var model = _ieService.GetOneComment(ie_id);
            return PartialView("_Comments", model);
        }

        #region POC method

        [HttpPost]
        public IActionResult GetPOC(int patientIEId, int patientFUId)
        {
            try
            {
                //var injurbodyparts = _pocService.GetInjuredParts(patientIEId, patientFUId);
                var injurbodyparts = _pocService.GetInjuredPartsPOC(patientIEId);


                if (injurbodyparts.Any(x => x.Equals("lumbar", StringComparison.OrdinalIgnoreCase)))
                {
                    injurbodyparts = injurbodyparts
                        .Where(x => !x.Equals("lowback", StringComparison.OrdinalIgnoreCase))
                        .ToArray();
                }
                if (injurbodyparts.Any(x => x.Equals("cervical", StringComparison.OrdinalIgnoreCase)))
                {
                    injurbodyparts = injurbodyparts
                        .Where(x => !x.Equals("neck", StringComparison.OrdinalIgnoreCase))
                        .ToArray();
                }
                if (injurbodyparts.Any(x => x.Equals("thoracic", StringComparison.OrdinalIgnoreCase)))
                {
                    injurbodyparts = injurbodyparts
                        .Where(x => !x.Equals("midback", StringComparison.OrdinalIgnoreCase))
                        .ToArray();
                }

                if (injurbodyparts != null)
                    injurbodyparts = injurbodyparts.Append("other").ToArray();
                else
                {
                    injurbodyparts = new string[1];
                    injurbodyparts[0] = "other";
                }

                string potion = null;
                string iinew = string.Empty;
                StringBuilder html = new StringBuilder();

                foreach (var ii in injurbodyparts)
                {
                    var _bodyparts = _commonservices.GetBodyPart(ii);
                    _bodyparts = _bodyparts.TrimStart();
                    _bodyparts = _bodyparts.TrimEnd();
                    if (_bodyparts.ToLower().Contains("left"))
                    {
                        potion = "Left";
                        iinew = _bodyparts.Substring(4, _bodyparts.Length - 4);
                    }
                    else if (_bodyparts.ToLower().Contains("right"))
                    {
                        potion = "Right";
                        iinew = _bodyparts.Substring(5, _bodyparts.Length - 5);
                    }
                    else
                    {
                        potion = null;
                        iinew = _bodyparts;
                    }
                    int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                    //var x = _pocService.GetAllProceduresFU(iinew.Trim(), patientFUId, potion, cmpid.Value); //commented by moulick as all poc required so specific visit wise poc cancled. 
                    var x = _pocService.GetAllProcedures(iinew.Trim(), patientIEId, potion, cmpid.Value);
                    //var x = _pocService.GetAllProcedures(iinew.Trim(), patientIEId, potion);

                    if (x != null)
                    {
                        if (x.Rows.Count > 0)
                        {

                            //Table start.
                            html.Append("<div class='accordion-item'>" +
                                "<h2 class='accordion-header' id='#" + ii + "_HeadingColps'>" +
                                "<button class='accordion-button collapsed' type='button' data-bs-toggle='collapse' data-bs-target='#" + ii.ToLower().Replace(" ", "_") + "_Colps' aria-expanded='false' aria-controls='flush-collapseTwo'>" + ii.ToUpperInvariant() + "</button></h2>" +
                                "<div id='" + ii.ToLower().Replace(" ", "_") + "_Colps' class='accordion-collapse collapse' aria-labelledby='" + ii.Replace(" ", "_") + "_HeadingColps' data-bs-parent='#accordionFlushExample'>" +
                                "<div class='accordion-body'><div class='table-responsive'>" +
                                "<table class='table table-hover table-bordered tblpoc'  border = '1' id='" + ii.ToLower().Replace(" ", "_") + "_tbl'>");
                            html.Append("<thead>");
                            //Building the Header row.
                            html.Append("<tr>");

                            foreach (DataColumn column in x.Columns)
                            {
                                if (column.ColumnName != "procedureID" && (column.ColumnName != "MCODE") &&
                                (column.ColumnName != "INhouseProc") && (column.ColumnName != "HasPosition") &&
                                (column.ColumnName != "HasLevel") && (column.ColumnName != "HasMuscle") &&
                                (column.ColumnName != "HasSubCode") && (column.ColumnName != "BID") && (column.ColumnName != "HasMedication") &&
                                (column.ColumnName != "PatientProceduresID")
                                && (column.ColumnName != "Medication") && (column.ColumnName != "Muscle") && (column.ColumnName != "Level")
                                && (column.ColumnName != "Req_Pos") && (column.ColumnName != "Sch_Pos") && (column.ColumnName != "FU_Pos")
                                && (column.ColumnName != "Exe_Pos") && (column.ColumnName != "SubProcedureID") && (column.ColumnName != "PatientIEID")
                                && (column.ColumnName != "PatientFuID") && (column.ColumnName != "Sides") && (column.ColumnName != "HasSides")
                                && (column.ColumnName != "Display_Order") && (column.ColumnName != "ProcedureDetailID")
                                && (column.ColumnName != "Consider") && (column.ColumnName != "Consider")
                                && (column.ColumnName != "LevelsDefault") && (column.ColumnName != "SidesDefault") && (column.ColumnName != "CF") && (column.ColumnName != "SignPath"))
                                {
                                    if (column.ColumnName == "bodypart")
                                    {
                                        html.Append("<th scope=\"col\" style='height: 35px;min-width:130px;background-color:yellow'>");
                                        html.Append(ii.ToUpperInvariant());
                                        html.Append("</th>");
                                    }
                                    else
                                    {

                                        if (column.ColumnName != "Followup" && column.ColumnName != "mcode_desc")
                                        {
                                            html.Append("<th scope=\"col\" style='height: 35px;min-width:130px'>");
                                            html.Append(column.ColumnName);
                                            html.Append("</th>");
                                        }
                                        /*else
                                        {
                                            html.Append("<th scope=\"col\"  style='height: 35px; background-color:grey'>");
                                            html.Append(column.ColumnName);
                                            html.Append("</th>");
                                        }*/

                                    }
                                }
                            }
                            html.Append("</tr>");
                            html.Append("</thead>");
                            html.Append("<tbody>");

                            foreach (DataRow row in x.Rows)
                            {
                                html.Append("<tr>");
                                foreach (DataColumn column in x.Columns)
                                {
                                    if ((column.ColumnName != "ProcedureDetail_ID") && (column.ColumnName != "Bodypart") && (column.ColumnName != "patientfu_id") && (column.ColumnName != "patientie_id") && (column.ColumnName != "InHouseProc"))
                                    {
                                        /*if (column.ColumnName == "Consider")
                                         {
                                             string date1 = string.Empty;
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[14])))
                                             {
                                                 date1 = Convert.ToDateTime(row[14]).ToString("MM/dd/yyyy").Replace('-', '/');
                                             }

                                             if (row[10] != null)
                                                 if (Convert.ToInt32(row[10]) != 0)
                                                 {
                                                     if (row[14] != DBNull.Value)
                                                     {
                                                         html.Append("<td >");
                                                         html.Append("<input type='button' class='btn btn-danger'   onclick='boxDisable($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "'   data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-Procedure_Detail_ID='" + row[30] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "'  data-Date='" + date1 + "' data-Position='Consider'   id='" + row[0] + "_C_" + row[1] + "_" + row[2] + "' value='Consider' />");
                                                         html.Append("</td>");
                                                     }
                                                     else
                                                     {
                                                         html.Append("<td >");
                                                         html.Append("<input type='button' class='btn btn-danger' onclick='boxDisable($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-Procedure_Detail_ID='" + row[30] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "'  data-Date='" + date1 + "' data-Position='Consider'   id='" + row[0] + "_C_" + row[1] + "_" + row[2] + "' value='Consider' />");
                                                         html.Append("</td>");
                                                     }
                                                 }
                                                 else
                                                 {
                                                     html.Append("<td >");
                                                     html.Append("<input type='button' class='btn btn-danger' onclick='boxDisable($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-Procedure_Detail_ID='" + row[30] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "'  data-Date='" + date1 + "' data-Position='Consider'   id='" + row[0] + "_C_" + row[1] + "_" + row[2] + "' value='Consider' />");
                                                     html.Append("</td>");
                                                 }
                                         }
                                         else*/
                                        if (column.ColumnName == "Requested")
                                        {

                                            string date1 = string.Empty;
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[15])))
                                            { date1 = Convert.ToDateTime(row[15]).ToString("MM/dd/yy").Replace('-', '/'); }

                                            StringBuilder notify = new StringBuilder();

                                            if (!string.IsNullOrEmpty(Convert.ToString(row[12])))
                                            { notify.AppendLine("muscle:" + row[12]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                            { notify.AppendLine("medication:" + row[11]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[23])))
                                            { notify.AppendLine("SubCode:" + row[23]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[27])))
                                            { notify.AppendLine("side:" + row[27]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[13])))
                                            { notify.AppendLine("level:" + row[13]); }
                                            if (row[10] != null)
                                                if (Convert.ToInt32(row[10]) != 0)
                                                {
                                                    html.Append("<td style='min-width:130px'>");
                                                    html.Append("<input type='text' class=\"form-control\"  onclick='PopupNE($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "'  data-toggle='tooltip' title='" + notify + "'  data-Sides='" + row[27] + "' data-Procedure_Detail_ID='" + row[30] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "'data-ReqPos='" + row[19] + "' data-Bodyid='" + row[7] + "' data-Position='Request' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'  class='ProcText' id='" + row[0] + "_R_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                                    html.Append("</td>");
                                                }
                                                else
                                                {
                                                    html.Append("<td style='min-width:130px;'>");
                                                    html.Append("<input type='text' class=\"form-control\" onclick='Popup($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "'  data-toggle='tooltip' title='" + notify + "'  data-Sides='" + row[27] + "' data-Procedure_Detail_ID='" + row[30] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Request' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'  class='ProcText' id='" + row[0] + "_R_" + row[1] + "_" + row[2] + "' value='' />");
                                                    html.Append("</td>");
                                                }
                                        }
                                        else if (column.ColumnName == "Scheduled")
                                        {

                                            string date1 = string.Empty;
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[16])))
                                            { date1 = Convert.ToDateTime(row[16]).ToString("MM/dd/yy").Replace('-', '/'); }


                                            StringBuilder notify = new StringBuilder();
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[12])))
                                            { notify.AppendLine("muscle:" + row[12]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                            { notify.AppendLine("medication:" + row[11]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[23])))
                                            { notify.AppendLine("SubCode:" + row[23]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[27])))
                                            { notify.AppendLine("side:" + row[27]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[13])))
                                            { notify.AppendLine("level:" + row[13]); }

                                            if (row[10] != null)
                                            {
                                                if (Convert.ToInt32(row[10]) != 0)
                                                {
                                                    html.Append("<td style='min-width:130px'>");
                                                    html.Append("<input type='text' class='ProcText form-control' onclick='PopupNE($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[30] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "'data-ReqPos='" + row[20] + "'data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Schedule' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "' data-MCode='" + row[1] + "'   id='" + row[0] + "_S_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                                    html.Append("</td>");
                                                }
                                                else
                                                {
                                                    html.Append("<td style='min-width:130px'>");
                                                    html.Append("<input type='text' class='ProcText form-control' onclick='Popup($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[30] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-PID='" + row[0] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Schedule' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "' data-MCode='" + row[1] + "'   id='" + row[0] + "_S_" + row[1] + "_" + row[2] + "' value='' />");
                                                    html.Append("</td>");
                                                }
                                            }

                                        }
                                        else if (column.ColumnName == "Executed")
                                        {

                                            string date1 = string.Empty;
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[17])))
                                            { date1 = Convert.ToDateTime(row[17]).ToString("MM/dd/yy").Replace('-', '/'); }


                                            StringBuilder notify = new StringBuilder();
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[12])))
                                            { notify.AppendLine("muscle:" + row[12]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                            { notify.AppendLine("medication:" + row[11]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[23])))
                                            { notify.AppendLine("SubCode:" + row[23]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[27])))
                                            { notify.AppendLine("side:" + row[27]); }
                                            if (!string.IsNullOrEmpty(Convert.ToString(row[13])))
                                            { notify.AppendLine("level:" + row[13]); }

                                            if (row[10] != null)
                                                if (Convert.ToInt32(row[10]) != 0)
                                                {
                                                    html.Append("<td style='min-width:130px'>");
                                                    html.Append("<input type='text' class='form-control dateval' onclick='PopupNE($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[30] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "'data-PID='" + row[0] + "'data-ReqPos='" + row[22] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Execute' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "' data-Medication='" + row[9] + "' data-SignPath='" + row[33] + "'    data-Date='" + date1 + "' data-MCode='" + row[1] + "'   class='ProcText' id='" + row[0] + "_E_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                                    html.Append("</td>");
                                                }
                                                else
                                                {
                                                    html.Append("<td style='min-width:130px'>");
                                                    html.Append("<input type='text' class='form-control' onclick='Popup($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[30] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-PID='" + row[0] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Execute' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "' data-Medication='" + row[9] + "'  data-Date='" + date1 + "' data-CF='" + row[34] + "' data-MCode='" + row[1] + "'     class='ProcText' id='" + row[0] + "_E_" + row[1] + "_" + row[2] + "' value='' />");
                                                    html.Append("</td>");
                                                }
                                        }
                                        /* else if (column.ColumnName == "Followup")
                                         {

                                             string date1 = string.Empty;
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[18])))
                                             { date1 = Convert.ToDateTime(row[18]).ToString("MM/dd/yyyy").Replace('-', '/'); }

                                             StringBuilder notify = new StringBuilder();
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[12])))
                                             { notify.AppendLine("muscle:" + row[12]); }
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                             { notify.AppendLine("medication:" + row[11]); }
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[23])))
                                             { notify.AppendLine("SubCode:" + row[23]); }
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[27])))
                                             { notify.AppendLine("side:" + row[27]); }
                                             if (!string.IsNullOrEmpty(Convert.ToString(row[13])))
                                             { notify.AppendLine("level:" + row[13]); }

                                             if (row[10] != null)
                                                 if (Convert.ToInt32(row[10]) != 0)
                                                 {
                                                     html.Append("<td style='background:grey'>");
                                                     html.Append("<input style='background:grey' type='text' class='form-control' onclick='PopupNE($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[30] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-PID='" + row[0] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "'data-ReqPos='" + row[21] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Follow Up' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'  class='ProcText' id='" + row[0] + "_F_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                                     html.Append("</td>");
                                                 }
                                                 else
                                                 {
                                                     html.Append("<td style='background:grey'>");
                                                     html.Append("<input style='background:grey' type='text' class='form-control' onclick='Popup($(this));' data-LevelsDefault='" + row[31] + "' data-SidesDefault='" + row[32] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[30] + "' data-Sides='" + row[27] + "' data-HasSides ='" + row[28] + "' data-PPID='" + row[10] + "' data-PatientIEID='" + row[24] + "' data-PatientFUID='" + row[25] + "' data-PID='" + row[0] + "' data-Level='" + row[13] + "' data-Medi='" + row[11] + "' data-Musc='" + row[12] + "' data-SubPID='" + row[23] + "' data-Body='" + row[8] + "' data-Bodyid='" + row[7] + "' data-Position='Follow Up' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'  class='ProcText' id='" + row[0] + "_F_" + row[1] + "_" + row[2] + "' value='' />");
                                                     html.Append("</td>");
                                                 }
                                         }*/
                                        else if (column.ColumnName == "count")
                                        {
                                            html.Append("<td style='min-width:130px'>");
                                            html.Append("<input type='button' class='btn btn-warning btn-sm' style='margin-left:25px' onclick='CountPopup($(this));' data-Div='" + ii.Replace(' ', '_') + "_counttable' data-Procedure_Detail_ID='" + row[30] + "' data-PatientIEID='" + row[24] + "' data-PID='" + row[0] + "' value='" + row[26] + "'  />");
                                            html.Append("</td>");
                                        }
                                        else if (column.ColumnName == "MCODE")
                                        {
                                            html.Append("<td  data-bs-toggle='tooltip' data-bs-placement='top'  title='" + row["mcode_desc"] + "' style='text-align:center; background-color:#3de33d;min-width:130px'>");
                                            html.Append(row[column.ColumnName]);
                                            html.Append("</td>");
                                        }
                                    }
                                }
                                html.Append("</tr>");
                            }
                            html.Append("</tbody></table></div><div id='" + ii.Replace(' ', '_') + "_counttable' style='padding-left: 1%' ></div></div></div></div>");

                            string body = html.ToString();

                            //Append the HTML string to Placeholder.
                            //PlaceHolder1.Controls.Add(new Literal { Text = html.ToString() });

                            //Page.ClientScript.RegisterStartupScript(this.GetType(), ii, "tableTransform($('#" + ii + "_tbl'));", true);

                        }
                        else
                        {
                            html.Append("<table border = '1' id='" + ii + "_tbl'>");
                            html.Append("</table>");
                            //Append the HTML string to Placeholder.
                            // PlaceHolder1.Controls.Add(new Literal { Text = html.ToString() });
                        }
                    }
                }

                ViewBag.content = html.ToString();

            }
            catch (Exception ex)
            {
                SaveLog(ex, "GetPOC");
            }
            return PartialView("_POC");
        }

        public string GetMuscleFromDB(int patientIEId)
        {
            string test = _pocService.GetMuscle(patientIEId);

            DataTable dt = new DataTable();
            string[] lines = test.Split('\n');
            dt.Columns.Add("Muscle");
            for (int i = 0; i < lines.Length; i++)
            {
                DataRow row = dt.NewRow();
                row[0] = lines[i];
                dt.Rows.Add(row);
            }
            string json = JsonConvert.SerializeObject(dt);
            return json;
        }

        public string GetSubCodeFromDB(int patientIEId)
        {
            DataTable dt = new DataTable();
            try
            {
                string test = _pocService.GetSubCode(patientIEId);
                string[] lines = test.Split('\n');
                dt.Columns.Add("SubCode");
                for (int i = 0; i < lines.Length; i++)
                {
                    DataRow row = dt.NewRow();
                    row[0] = lines[i];
                    dt.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "GetSubCodeFromDB");
            }
            string json = JsonConvert.SerializeObject(dt);
            return json;
        }

        public string GetMedicationFromDB(int patientIEId)
        {
            DataTable dt = new DataTable();
            try
            {
                string test = _pocService.GetMedication(patientIEId);

                string[] lines = test.Split('\n');
                dt.Columns.Add("Medicaton");
                for (int i = 0; i < lines.Length; i++)
                {
                    DataRow row = dt.NewRow();
                    row[0] = lines[i];
                    dt.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex, "GetMedicationFromDB");
            }
            string json = JsonConvert.SerializeObject(dt);
            return json;
        }

        [HttpPost]
        public JsonResult SavePOC(ProcedureDetailsVMNew model)
        {
            string Req_Pos = ""; string Sch_Pos = ""; string Exe_Pos = ""; string FU_Pos = "", result = "0";

            string count = "", bodyPartSide = "";

            if (!string.IsNullOrEmpty(model.Position))
                bodyPartSide = model.Position + " " + model.BodyPart;
            else
                bodyPartSide = model.BodyPart;

            try
            {
                var data = new ProcedureDetailsVM();

                data.Backup_Line = model.Backup_Line;
                data.BodyPart = model.BodyPart;
                data.BodypartSide = bodyPartSide;
                data.CreatedBy = "";
                data.CT_AUTH_Date = Common.ConvertDate(model.CT_AUTH_Date);
                data.CT_Note = model.CT_Note;
                data.CT_Report_Status = model.CT_Report_Status;
                data.CT_ReSche_Date = Common.ConvertDate(model.CT_ReSche_Date);
                data.Ins_Note = model.Ins_Note;
                data.Ins_Ver_Status = Convert.ToBoolean(model.Ins_Ver_Status);
                data.IsConsidered = model.IsConsidered == "0" ? false : true;
                data.IsFromNew = Convert.ToInt16(model.IsFromNew);
                data.IsVaccinated = Convert.ToBoolean(model.IsVaccinated);
                data.Level = model.Level;
                data.MC_Date = Common.ConvertDate(model.MC_Date);
                data.MC_Note = model.MC_Note;
                data.MC_Report_Status = model.MC_Report_Status;
                data.MC_ReSche_Date = Common.ConvertDate(model.MC_ReSche_Date);
                data.MC_Type = model.MC_Type;
                data.Medication = model.Medication;
                data.Muscle = model.Muscle;
                data.ProcedureMasterID = string.IsNullOrEmpty(model.ProcedureMasterID) ? null : Convert.ToInt32(model.ProcedureMasterID);
                data.PatientFuID = string.IsNullOrEmpty(model.PatientFuID) ? null : Convert.ToInt32(model.PatientFuID);
                data.PatientIEID = string.IsNullOrEmpty(model.PatientIEID) ? null : Convert.ToInt32(model.PatientIEID);
                data.PatientProceduresID = string.IsNullOrEmpty(model.PatientProcedureID) ? null : Convert.ToInt32(model.PatientProcedureID);
                data.ProcedureDetailID = string.IsNullOrEmpty(model.ProcedureDetailID) ? null : Convert.ToInt32(model.ProcedureDetailID);
                data.ProcedureID = string.IsNullOrEmpty(model.ProcedureID) ? null : Convert.ToInt32(model.ProcedureID);
                data.Side = model.Side;
                data.SignPath = "";
                data.Subcode = model.SubProcedureID;
                data.Vac_Note = model.Vac_Note;
                data.Vac_Status = model.Vac_Status;
                data.Cmp_Id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                switch (model.req)
                {
                    case "Request":
                        DateTime? Requested = Common.ConvertDate(model.pocdate);
                        data.Req_Pos = model.Position;
                        data.Requested = Requested;
                        result = _pocService.SaveProcedureDetails(data);
                        // count = _bl.savePatientProcedureDetailNew(ProcedureDetailID, ProcedureMasterID, _patientIEID, _patientFUID, BodyPartID, ProcedureID, null, Requested, null, null, null, BodyPart, Medication, Muscle, Level, Req_Pos, Sch_Pos, Exe_Pos, FU_Pos, 0, IsFromNew, PatientProcedureID, req, IsConsidered, Side, SubProcedureID, null, bodyPartSide, MC_Type, MC_Date, MC_Report_Status, MC_Note, MC_ReSche_Date, CT_AUTH_Date, CT_Report_Status, CT_Note, CT_ReSche_Date, Ins_Ver_Status, Backup_Line, Ins_Note, IsVaccinated, Vac_Status, Vac_Note);
                        break;
                    case "Schedule":
                        DateTime? Scheduled = Common.ConvertDate(model.pocdate);
                        data.Sch_Pos = model.Position;
                        data.Scheduled = Scheduled;
                        result = _pocService.SaveProcedureDetails(data);
                        //count = _bl.savePatientProcedureDetailNew(ProcedureDetailID, ProcedureMasterID, _patientIEID, _patientFUID, BodyPartID, ProcedureID, null, null, Scheduled, null, null, BodyPart, Medication, Muscle, Level, Req_Pos, Sch_Pos, Exe_Pos, FU_Pos, 0, IsFromNew, PatientProcedureID, req, IsConsidered, Side, SubProcedureID, null, bodyPartSide, MC_Type, MC_Date, MC_Report_Status, MC_Note, MC_ReSche_Date, CT_AUTH_Date, CT_Report_Status, CT_Note, CT_ReSche_Date, Ins_Ver_Status, Backup_Line, Ins_Note, IsVaccinated, Vac_Status, Vac_Note);
                        break;

                    case "Execute":
                        DateTime? Executed = Common.ConvertDate(model.pocdate);
                        data.Exe_Pos = model.Position;
                        data.Executed = Executed;
                        result = _pocService.SaveProcedureDetails(data);
                        break;
                }


                int fu_id = Convert.ToInt32(model.PatientFuID);
                var pocData = this.UpdatePOCPlan(fu_id);

                if (pocData != null)
                {
                    //var obj = _fuPage1services.GetOne(fu_id);
                    //string strCCdesc = "", strPEdesc = "", strAdesc = "";
                    //if (obj != null)
                    //{
                    //    if (!string.IsNullOrEmpty(pocData.strCCDesc))
                    //        strCCdesc = obj.cc + "<p>" + pocData.strCCDesc + "</p>";
                    //    else
                    //        strCCdesc = obj.cc;

                    //    if (!string.IsNullOrEmpty(pocData.strPEDesc))
                    //        strPEdesc = obj.pe + "<p>" + pocData.strPEDesc + "</p>";
                    //    else
                    //        strPEdesc = obj.pe;

                    //    if (!string.IsNullOrEmpty(pocData.strADesc))
                    //        strAdesc = obj.assessment + "<p>" + pocData.strADesc + "</p>";
                    //    else
                    //        strAdesc = obj.assessment;
                    //}
                    //pocData.strADesc = strAdesc;
                    //pocData.strCCDesc = strCCdesc;
                    //pocData.strPEDesc = strPEdesc;
                    _patientFuservices.UpdatePage1Plan(fu_id, pocData.strPoc);
                }
                return Json(pocData);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SavePOC");
            }
            //_pocService.SaveProcedureDetails(model);
            return Json(result);
        }

        [HttpPost]
        public string GetProcDetail(int ProcedureDetailID)
        {
            var data = _pocService.GetProcedureDetails(ProcedureDetailID);
            string json = JsonConvert.SerializeObject(data);
            return json;
        }

        [HttpPost]
        public string GetProcedureCountDetail(int ProcedureID, int PatientIEID)
        {
            var x = _pocService.GetProcedureCountDetails(ProcedureID, PatientIEID);
            StringBuilder html = new StringBuilder();
            if (x.Rows.Count > 0)
            {
                // x = FlipDataTable(x);

                //Table start.
                html.Append("<table class='table table-hover table-bordered tblpoc'  border = '1' id='countbl'>");
                html.Append("<thead>");
                //Building the Header row.
                html.Append("<tr>");
                foreach (DataColumn column in x.Columns)
                {

                    if (column.ColumnName != "procedureID" && (column.ColumnName != "MCODE") &&
                    (column.ColumnName != "INhouseProc") && (column.ColumnName != "Position") &&
                    (column.ColumnName != "HasLevel") && (column.ColumnName != "HasMuscle") &&
                    (column.ColumnName != "HasMedication") && (column.ColumnName != "PatientProceduresID")
                    && (column.ColumnName != "Medication") && (column.ColumnName != "Muscle") && (column.ColumnName != "Level")
                    && (column.ColumnName != "Req_Pos") && (column.ColumnName != "Sch_Pos") && (column.ColumnName != "FU_Pos")
                    && (column.ColumnName != "Exe_Pos") && (column.ColumnName != "PatientIEID") && (column.ColumnName != "PatientFuID")
                    && (column.ColumnName != "SubProcedureID") && (column.ColumnName != "HasSubCode")
                    && (column.ColumnName != "HasSides") && (column.ColumnName != "sides") && (column.ColumnName != "ProcedureDetailID")
                    && (column.ColumnName != "LevelsDefault") && (column.ColumnName != "SidesDefault"))
                    {
                        if (column.ColumnName == "BodyPart")
                        {
                            html.Append("<th style='height: 35px; background-color:yellow'>");
                            html.Append("");
                            html.Append("</th>");
                        }
                        else
                        {
                            if (column.ColumnName == "Consider" || column.ColumnName == "Followup")
                            {
                                html.Append("<th style='display:none'>");
                                html.Append(column.ColumnName);
                                html.Append("</th>");
                            }
                            else if (column.ColumnName == "Followup")
                            {
                                html.Append("<th style='height: 35px;background:grey'>");
                                html.Append(column.ColumnName);
                                html.Append("</th>");
                            }
                            else
                            {
                                html.Append("<th  style='height: 35px;'>");
                                html.Append(column.ColumnName);
                                html.Append("</th>");
                            }

                        }
                    }
                    //}
                }
                html.Append("</tr>");
                html.Append("</thead>");
                html.Append("<tbody>");

                foreach (DataRow row in x.Rows)
                {
                    html.Append("<tr>");
                    foreach (DataColumn column in x.Columns)
                    {
                        if ((column.ColumnName != "PatientProceduresID") && (column.ColumnName != "BodyPart") && (column.ColumnName != "PatientFuID") && (column.ColumnName != "PatientIEID") && (column.ColumnName != "INhouseProc"))
                        {
                            //|| column.ColumnName == "Consider" || )
                            if (column.ColumnName == "Consider")
                            {
                                if (row[8] != null)
                                    if (Convert.ToInt32(row[8]) != 0)
                                    {
                                        if (row[12] != DBNull.Value)
                                        {
                                            html.Append("<td style='display:none'>");
                                            html.Append("<div class='checkbox'><input type='checkbox' checked  onclick='boxDisable($(this));' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "' data-Body='" + row[6] + "' data-Date='" + row[12] + "' data-Position='Consider'   id='" + row[0] + "_C_" + row[1] + "_" + row[2] + "' value='Consider' /></div>");
                                            html.Append("</td>");
                                        }
                                        else
                                        {
                                            html.Append("<td style='display:none'>");
                                            html.Append("<div class='checkbox'><input type='checkbox' onclick='boxDisable($(this));' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "' data-Body='" + row[6] + "' data-Date='" + row[12] + "' data-Position='Consider'   id='" + row[0] + "_C_" + row[1] + "_" + row[2] + "' value='Consider' /></div>");
                                            html.Append("</td>");
                                        }
                                    }
                                    else
                                    {
                                        html.Append("<td style='display:none'>");
                                        html.Append("<div class='checkbox'><input type='checkbox' onclick='boxDisable($(this));' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "' data-Body='" + row[6] + "' data-Date='" + row[12] + "' data-Position='Consider'   id='" + row[0] + "_C_" + row[1] + "_" + row[2] + "' value='Consider' /></div>");
                                        html.Append("</td>");
                                    }
                            }
                            else if (column.ColumnName == "Requested")
                            {
                                string date1 = string.Empty;
                                if (!string.IsNullOrEmpty(Convert.ToString(row[13])))
                                { date1 = Convert.ToDateTime(row[13]).ToString("MM/dd/yyyy").Replace('-', '/'); }


                                StringBuilder notify = new StringBuilder();
                                if (!string.IsNullOrEmpty(Convert.ToString(row[10])))
                                { notify.AppendLine("muscle:" + row[10]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[9])))
                                { notify.AppendLine("medication:" + row[9]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[24])))
                                { notify.AppendLine("SubCode:" + row[24]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[26])))
                                { notify.AppendLine("side:" + row[26]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                { notify.AppendLine("level:" + row[11]); }

                                if (row[8] != null)
                                    if (Convert.ToInt32(row[8]) != 0)
                                    {
                                        html.Append("<td >");
                                        html.Append("<input type='text'  class=\"form-control\" data-toggle='tooltip' title='" + notify + "' onclick='PopupNE($(this));' data-LevelsDefault='" + row[28] + "' data-SidesDefault='" + row[29] + "' data-PPID='" + row[8] + "' data-Procedure_Detail_ID='" + row[27] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "' data-Body='" + row[6] + "'data-ReqPos='" + row[17] + "' data-Position='Request' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'  class='ProcText' id='" + row[0] + "_R_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                        html.Append("</td>");
                                    }
                                    else
                                    {
                                        html.Append("<td >");
                                        html.Append("<input type='text'  class=\"form-control\" onclick='Popup($(this));'  data-LevelsDefault='" + row[28] + "' data-SidesDefault='" + row[29] + "' data-toggle='tooltip' title='" + notify + "'  data-PPID='" + row[8] + "'  data-Procedure_Detail_ID='" + row[27] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "' data-Body='" + row[6] + "' data-Position='Request' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'  class='ProcText' id='" + row[0] + "_R_" + row[1] + "_" + row[2] + "' value='' />");
                                        html.Append("</td>");
                                    }
                            }
                            else if (column.ColumnName == "Scheduled")
                            {

                                string date1 = string.Empty;
                                if (!string.IsNullOrEmpty(Convert.ToString(row[14])))
                                { date1 = Convert.ToDateTime(row[14]).ToString("MM/dd/yyyy").Replace('-', '/'); }

                                StringBuilder notify = new StringBuilder();
                                if (!string.IsNullOrEmpty(Convert.ToString(row[10])))
                                { notify.AppendLine("muscle:" + row[10]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[9])))
                                { notify.AppendLine("medication:" + row[9]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[24])))
                                { notify.AppendLine("SubCode:" + row[24]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[26])))
                                { notify.AppendLine("side:" + row[26]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                { notify.AppendLine("level:" + row[11]); }
                                if (row[8] != null)
                                {
                                    if (Convert.ToInt32(row[8]) != 0)
                                    {
                                        html.Append("<td >");
                                        html.Append("<input type='text'  class=\"form-control\" class='ProcText' data-toggle='tooltip' title='" + notify + "'  onclick='PopupNE($(this));' data-LevelsDefault='" + row[28] + "' data-SidesDefault='" + row[29] + "' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "'data-ReqPos='" + row[18] + "' data-Body='" + row[6] + "' data-Position='Schedule' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'   id='" + row[0] + "_S_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                        html.Append("</td>");
                                    }
                                    else
                                    {
                                        html.Append("<td >");
                                        html.Append("<input type='text'  class=\"form-control\" class='ProcText' data-toggle='tooltip' title='" + notify + "' onclick='Popup($(this));' data-LevelsDefault='" + row[28] + "' data-SidesDefault='" + row[29] + "' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-PID='" + row[0] + "' data-Body='" + row[6] + "' data-Position='Schedule' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-InhouseProc='" + row[2] + "'  data-Medication='" + row[9] + "'  data-Date='" + date1 + "'   id='" + row[0] + "_S_" + row[1] + "_" + row[2] + "' value='' />");
                                        html.Append("</td>");
                                    }
                                }

                            }
                            else if (column.ColumnName == "Executed")
                            {
                                string date1 = string.Empty;
                                if (!string.IsNullOrEmpty(Convert.ToString(row[15])))
                                { date1 = Convert.ToDateTime(row[15]).ToString("MM/dd/yyyy").Replace('-', '/'); }

                                StringBuilder notify = new StringBuilder();
                                if (!string.IsNullOrEmpty(Convert.ToString(row[10])))
                                { notify.AppendLine("muscle:" + row[10]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[9])))
                                { notify.AppendLine("medication:" + row[9]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[24])))
                                { notify.AppendLine("SubCode:" + row[24]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[26])))
                                { notify.AppendLine("side:" + row[26]); }
                                if (!string.IsNullOrEmpty(Convert.ToString(row[11])))
                                { notify.AppendLine("level:" + row[11]); }

                                if (row[8] != null)
                                    if (Convert.ToInt32(row[8]) != 0)
                                    {
                                        html.Append("<td >");
                                        html.Append("<input type='text'  class=\"form-control\" onclick='PopupNE($(this));' data-LevelsDefault='" + row[28] + "' data-SidesDefault='" + row[29] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "'data-PID='" + row[0] + "'data-ReqPos='" + row[20] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-Body='" + row[6] + "' data-Position='Execute' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-InhouseProc='" + row[2] + "' data-Medication='" + row[9] + "'  data-Date='" + date1 + "' data-MCode='" + row[1] + "'   class='ProcText' id='" + row[0] + "_E_" + row[1] + "_" + row[2] + "' value='" + date1 + "' />");
                                        html.Append("</td>");
                                    }
                                    else
                                    {
                                        html.Append("<td >");
                                        html.Append("<input type='text'  class=\"form-control\" onclick='Popup($(this));' data-LevelsDefault='" + row[28] + "' data-SidesDefault='" + row[29] + "' data-toggle='tooltip' title='" + notify + "' data-Procedure_Detail_ID='" + row[27] + "' data-PPID='" + row[8] + "' data-PatientIEID='" + row[21] + "' data-PatientFUID='" + row[22] + "' data-PID='" + row[0] + "' data-Level='" + row[11] + "' data-Medi='" + row[9] + "' data-Musc='" + row[10] + "' data-Body='" + row[6] + "' data-Position='Execute' data-HasLevel='" + row[4] + "' data-Pos='" + row[3] + "' data-Muscle='" + row[5] + "' data-Subcode='" + row[6] + "' data-InhouseProc='" + row[2] + "' data-Medication='" + row[9] + "'  data-Date='" + date1 + "' data-MCode='" + row[1] + "'   class='ProcText' id='" + row[0] + "_E_" + row[1] + "_" + row[2] + "' value='' />");
                                        html.Append("</td>");
                                    }
                            }

                            else if (column.ColumnName == "MCODE")
                            {
                                html.Append("<td style='text-align:center; background-color:#3de33d'>");
                                html.Append(row[column.ColumnName]);
                                html.Append("</td>");
                            }
                        }
                    }
                    html.Append("</tr>");
                }
                html.Append("</tbody>");
                //Table end.
                html.Append("</table>");

            }
            return html.ToString();
        }
        [HttpPost]
        public IActionResult Delete(int ProcedureDetailID, int fu_id)
        {
            var data = _pocService.RemoveProcedureCountDetails(ProcedureDetailID);

            var pocData = this.UpdatePOCPlan(fu_id);

            if (pocData != null)
            {
                var obj = _fuPage1services.GetOne(fu_id);
                string strCCdesc = "", strPEdesc = "", strAdesc = "";
                if (obj != null)
                {
                    if (!string.IsNullOrEmpty(pocData.strCCDesc))
                        strCCdesc = obj.cc + "<p>" + pocData.strCCDesc + "</p>";
                    else
                        strCCdesc = obj.cc;

                    if (!string.IsNullOrEmpty(pocData.strPEDesc))
                        strPEdesc = obj.pe + "<p>" + pocData.strPEDesc + "</p>";
                    else
                        strPEdesc = obj.pe;

                    if (!string.IsNullOrEmpty(pocData.strADesc))
                        strAdesc = obj.assessment + "<p>" + pocData.strADesc + "</p>";
                    else
                        strAdesc = obj.assessment;
                }
                pocData.strADesc = strAdesc;
                pocData.strCCDesc = strCCdesc;
                pocData.strPEDesc = strPEdesc;
                _patientFuservices.UpdatePage1Plan(fu_id, pocData.strPoc);
            }
            return Json(pocData.strPoc);
        }

        [HttpPost]
        public IActionResult DeleteFU(int fuId, string pName)
        {
            try
            {
                _patientFuservices.Delete(fuId);
                string cmpid = HttpContext.Session.GetString(SessionKeys.SessionCmpClientId);
                string msg = "FUID: " + fuId + " and PatientName: " + pName + " is deleted by " + HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId).ToString() + " in company id: " + cmpid;
                var logdata = new tbl_log
                {
                    CreatedDate = DateTime.Now,
                    CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId),
                    Message = msg
                };
                new LogService().Insert(logdata);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "DeleteFU");
            }
            return Json(1);
        }


        [HttpPost]
        public IActionResult GetDaignoCodeList(string bodyparts, int fuId)
        {

            var page1Data = _fuPage1services.GetOne(fuId);

            string assetment = "";

            if (page1Data != null)
                assetment = page1Data.assessment;


            bodyparts = bodyparts.Replace("_", " ");
            bodyparts = bodyparts.TrimEnd();
            ViewBag.BodyPart = bodyparts.ToUpper();
            var _bodyparts = _commonservices.GetBodyPart(bodyparts);
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

            string cnd = " and cmp_id=" + cmpid + " and (BodyPart='" + _bodyparts + "' or Description like '%" + _bodyparts + "%') order by Description desc";

            var data = _diagcodesService.GetAll(cnd);


            var datavm = (from c in data
                          select new DaignoCodeVM
                          {
                              DaignoCodeId = c.Id.Value,
                              Description = c.Description,
                              DiagCode = c.DiagCode,
                              IsSelect = assetment != null ? (assetment.IndexOf(c.Description) > 0 ? true : c.PreSelect) : c.PreSelect,
                              Display_Order = c.display_order,
                              cmp_id = c.cmp_id

                          }).ToList().Where(x => x.cmp_id.Value.ToString() == cmpid).OrderBy(x => x.Display_Order);
            return PartialView("_DaignoCode", datavm);
        }

        #endregion

        #region POCSumarry
        [HttpPost]
        public IActionResult POCSummary(int patientFUId)
        {
            try
            {
                var Data = _pocService.GetFUPOCSummary(patientFUId);
                return PartialView("_POCSummary", Data);

            }
            catch (Exception ex)
            {
                SaveLog(ex, "POCSummary");
                return View();
            }

        }

        [HttpPost]
        public IActionResult SurgeryPOCSummary(int patientIEId)
        {
            var Data = _pocService.GetExecutedPOCIE(patientIEId);
            try
            {
                int? ieId = HttpContext.Session.GetInt32(SessionKeys.SessionIEId);
            }
            catch (Exception ex)
            {
                SaveLog(ex, "SurgeryPOCSummary");

            }
            return PartialView("_POCSummary", Data);
        }

        [HttpGet]
        public JsonResult GetPhysiciansByLocation(int locationId)
        {
            var physicians = _physicianService.GetAll(" AND locationid=" + locationId);
            var physicianList = physicians.Select(p => new SelectListItem
            {
                Text = p.physicianname,
                Value = p.Id.ToString(),
            }).ToList();

            return Json(physicianList);
        }
        #endregion

        #region PreOpPrint
        public IActionResult PreOpPrint(int ieid, int fuid)
        {
            try
            {
                string plan = "";
                ViewBag.url = HttpContext.Request.GetEncodedUrl();

                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                string body = string.Empty;

                //using streamreader for reading my htmltemplate   

                var fuData = _patientFuservices.GetOne(fuid);

                tbl_users user = new tbl_users();



                var templateData = _printService.GetTemplate(cmpid, fuData.type);
                var gender = "";

                body = templateData.content;

                var patientData = _ieService.GetOnebyPatientId(ieid);
                body = body.Replace("#Physicianhistory", patientData.providerName);

                if (patientData != null)
                {
                    gender = Common.GetMrMrsFromSex(patientData.gender);

                    body = body.Replace("#patientname", gender + " " + patientData.fname + " " + patientData.mname + " " + patientData.lname);
                    body = body.Replace("#fn", patientData.fname);
                    body = body.Replace("#ln", patientData.lname);
                    body = body.Replace("#dob", Common.commonDate(patientData.dob));
                    body = body.Replace("#doi", Common.commonDate(patientData.doa));
                    body = body.Replace("#dos", Common.commonDate(fuData.doe));
                    //body = body.Replace("#location", patientData.location);
                    body = body.Replace("#age", patientData.age == null ? "0" : patientData.age.Value.ToString());
                    body = body.Replace("#upper_gender", Common.GetGenderFromSex(patientData.gender).First().ToString().ToUpper() + Common.GetGenderFromSex(patientData.gender).Substring(1));
                    body = body.Replace("#gender", Common.GetGenderFromSex(patientData.gender));
                    body = body.Replace("#sex", Common.GetGenderFromSex(patientData.gender));
                    body = body.Replace("#acctno", patientData.account_no);
                    body = body.Replace("#CASETYPE", patientData.accidentType);
                }

                //header printing

                var locData = _locService.GetAll(" and id=" + fuData.location_id);

                if (locData != null && locData.Count > 0)
                {
                    if (locData[0].nameofpractice != null)
                        body = body.Replace("#drFLName", locData[0].drfname + " " + locData[0].drlname);
                    else
                        body = body.Replace("#drFLName", "");
                    body = body.Replace("#drName", locData[0].nameofpractice.ToLower().Contains("dr") ? locData[0].nameofpractice : locData[0].nameofpractice);
                    body = body.Replace("#address", locData[0].address);
                    //body = body.Replace("#Address", locData[0].address);
                    body = body.Replace("#Address", locData[0].address + "<br/>" + locData[0].city + ", " + locData[0].state + " " + locData[0].zipcode);
                    body = body.Replace("#location", locData[0].location);
                    body = body.Replace("#Nameofpractice", locData[0].nameofpractice.ToLower().Contains("dr") ? locData[0].nameofpractice : locData[0].nameofpractice);
                    body = body.Replace("#Phone", locData[0].telephone);
                }
                //Preop printing

                var preData = _fuPreservices.GetOne(fuid);
                string bodypart = "";
                var page1Data = _ieService.GetOnePage1(ieid);
                if (page1Data != null)
                {
                    if (!string.IsNullOrEmpty(page1Data.bodypart))
                    {
                        bodypart = Common.ReplceCommowithAnd(page1Data.bodypart.ToLower());
                        body = body.Replace("#PC", string.IsNullOrEmpty(bodypart) ? "" : bodypart.Replace(",", ", ") + " pain.");

                    }
                    else
                    {
                        body = body.Replace("#PC", "");
                    }
                }

                var Presentillness = string.IsNullOrEmpty(preData.txtPresentillness) ? "" : preData.txtPresentillness;
                Presentillness = Presentillness.Replace("#age", patientData.age == null ? "0" : patientData.age.Value.ToString());
                Presentillness = Presentillness.Replace("#sex", Common.GetGenderFromSex(patientData.gender));
                Presentillness = Presentillness.Replace("#PC", string.IsNullOrEmpty(bodypart) ? "" : bodypart.Replace(",", ", ") + " pain.");
                Presentillness = Presentillness.Replace("#CASETYPE", patientData.accidentType);
                Presentillness = Presentillness.Replace("#accidenttype", patientData.accidentType);
                Presentillness = Presentillness.Replace("#doi", Common.commonDate(patientData.doa));
                //Presentillness = Presentillness.Replace("#ProviderName", Common.commonDate(patientData.doa));

                body = body.Replace("#Presentillness", Presentillness);
                var strDiagnostic = this.getDiagnosticie(ieid, preData);

                if (cmpid != "4")
                {
                    if (string.IsNullOrEmpty(strDiagnostic))
                    {
                        if (HttpContext.Session.GetString(SessionKeys.SessionIsDaignosis) == "true")
                        {
                            strDiagnostic = HttpContext.Session.GetString(SessionKeys.SessionDaignosisNotFoundStatment);
                        }
                    }
                    else
                    {
                        if (HttpContext.Session.GetString(SessionKeys.SessionIsDaignosis) == "true")
                        {
                            strDiagnostic = strDiagnostic + "<br/><br/>" + HttpContext.Session.GetString(SessionKeys.SessionDaignosisFoundStatment); ;
                        }
                    }
                }

                body = body.Replace("#Diagnostic", this.removePtag(strDiagnostic));
                if (preData != null)
                {
                    body = body.Replace("#CC", string.IsNullOrEmpty(preData.txtHistoryPresentillness) ? "" : this.removePtag(preData.txtHistoryPresentillness));
                    body = body.Replace("#presentcomplain", string.IsNullOrEmpty(preData.txtpresentcomplain) ? "" : this.removePtag(preData.txtpresentcomplain));
                    body = body.Replace("#PastMedicalHistory", string.IsNullOrEmpty(preData.txtPastMedicalHistory) ? "" : this.removePtag(preData.txtPastMedicalHistory));
                    body = body.Replace("#PastSurgicalHistory", string.IsNullOrEmpty(preData.txtpastsurgicalhistory) ? "" : this.removePtag(preData.txtpastsurgicalhistory));
                    body = body.Replace("#CURRENTMEDICATIONS", string.IsNullOrEmpty(preData.txtdailyMedications) ? "" : this.removePtag(preData.txtdailyMedications));
                    body = body.Replace("#Allergies", string.IsNullOrEmpty(preData.txtAllergies) ? "" : this.removePtag(preData.txtAllergies));
                    body = body.Replace("#SocialHistory", string.IsNullOrEmpty(preData.txtSH) ? "" : this.removePtag(preData.txtSH));
                    body = body.Replace("#FamilyHistory", string.IsNullOrEmpty(preData.txtFamilyHistory) ? "" : this.removePtag(preData.txtFamilyHistory));
                    body = body.Replace("#REVIEWOFSYSTEMS", string.IsNullOrEmpty(preData.txtSocialHistory) ? "" : this.removePtag(preData.txtSocialHistory));
                    body = body.Replace("#PhysicalExamination", string.IsNullOrEmpty(preData.txtPhysicalExamination) ? "" : this.removePtag(preData.txtPhysicalExamination));
                    body = body.Replace("#DIAGNOSTICSTUDIES", string.IsNullOrEmpty(preData.txtDiagnosticImaging) ? "" : this.removePtag(preData.txtDiagnosticImaging));
                    body = body.Replace("#DISABILITY", string.IsNullOrEmpty(preData.txtAssestmentplan) ? "" : this.removePtag(preData.txtAssestmentplan));
                    body = body.Replace("#PROPOSEDTREATMENT", string.IsNullOrEmpty(preData.txtExaminedResult) ? "" : this.removePtag(preData.txtExaminedResult));
                    body = body.Replace("#Default", string.IsNullOrEmpty(preData.txtDefault) ? "" : this.removePtag(preData.txtDefault));
                    body = body.Replace("#Note", string.IsNullOrEmpty(preData.txtNote) ? "" : this.removePtag(preData.txtNote));
                    ViewBag.ieId = patientData.id;
                    ViewBag.fuId = fuid;
                    ViewBag.locId = patientData.location_id;


                    string signName = "";
                    int signUserId = 0;

                    int? providorId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId);

                    var providerData = new tbl_users();
                    if (fuData.provider_id != null)
                    {
                        user.Id = fuData.provider_id;
                        providerData = _userService.GetOneById(user.Id.Value);
                        if (providerData.Id != null)
                        {
                            signUserId = providerData.Id.Value;
                        }
                        else if (providorId != null)
                        {
                            signUserId = providorId.Value;
                        }
                    }




                    if (signUserId > 0)
                    {
                        tbl_users _user = new tbl_users()
                        {
                            Id = signUserId
                        };
                        var userData = _userService.GetOne(_user);
                        signName = userData.signature;

                        if (!string.IsNullOrEmpty(signName))
                        {
                            string signatureUrl = $"/Uploads/Sign/" + cmpid + "/" + signName;
                            //string signatureUrl = "https://paintrax.com/newversionlive/Uploads/Sign/" + cmpid + "/" + signName;
                            string base64Image = ImageToBase64(Environment.WebRootPath + signatureUrl);
                            body = body.Replace("#Sign", $" <img src='data:image/jpg;base64,{base64Image}' alt='My Image' />");
                            // body = body.Replace("#Sign", $"<img crossorigin='anonymous|use-credentials' src='{signatureUrl}' alt='Patient Signature' />");
                        }
                        else
                            body = body.Replace("#Sign", "");

                        body = body.Replace("#Physician", providerData.providername);

                        body = body.Replace("#ProviderName", providerData.providername);
                        body = body.Replace("#AssProviderName", providerData.assistant_providername);
                    }
                    else
                        body = body.Replace("#Sign", "");
                }
                ViewBag.content = body;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "PreOpPrint");
            }
            return View();
        }
        #endregion
        #region PostOpPrint
        public IActionResult PostOpPrint(int ieid, int fuid)
        {
            try
            {
                string plan = "";
                ViewBag.url = HttpContext.Request.GetEncodedUrl();

                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                string body = string.Empty;

                //using streamreader for reading my htmltemplate   

                var fuData = _patientFuservices.GetOne(fuid);

                tbl_users user = new tbl_users();
                user.Id = fuData.provider_id;
                var providerData = _userService.GetOneById(user.Id.Value);


                var templateData = _printService.GetTemplate(cmpid, fuData.type);
                var gender = "";
                var page1Data = _fuPage1services.GetOne(fuid);
                var postData = _fuPostService.GetOne(fuid);
                body = templateData.content;
                var data = getPOCDate(fuid, ieid);
                var patientData = _ieService.GetOnebyPatientId(ieid);
                body = body.Replace("#Physicianhistory", patientData.providerName);

                if (patientData != null)
                {
                    gender = Common.GetMrMrsFromSex(patientData.gender);

                    body = body.Replace("#patientname", gender + " " + patientData.fname + " " + patientData.mname + " " + patientData.lname);
                    body = body.Replace("#fn", patientData.fname);
                    body = body.Replace("#ln", patientData.lname);
                    body = body.Replace("#dob", Common.commonDate(patientData.dob));
                    body = body.Replace("#doi", Common.commonDate(patientData.doa));
                    body = body.Replace("#dos", Common.commonDate(Convert.ToDateTime(fuData.doe), HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));
                    //body = body.Replace("#location", patientData.location);
                    body = body.Replace("#age", patientData.age == null ? "0" : patientData.age.Value.ToString());
                    body = body.Replace("#upper_gender", Common.GetGenderFromSex(patientData.gender).First().ToString().ToUpper() + Common.GetGenderFromSex(patientData.gender).Substring(1));
                    body = body.Replace("#gender", Common.GetGenderFromSex(patientData.gender));
                    body = body.Replace("#sex", Common.GetGenderFromSex(patientData.gender));
                    body = body.Replace("#acctno", patientData.account_no);
                    body = body.Replace("#CASETYPE", patientData.accidentType);
                }

                //header printing

                var locData = _locService.GetAll(" and id=" + fuData.location_id);

                if (locData != null && locData.Count > 0)
                {
                    if (locData[0].nameofpractice != null)
                        body = body.Replace("#drFLName", locData[0].drfname + " " + locData[0].drlname);
                    else
                        body = body.Replace("#drFLName", "");
                    body = body.Replace("#drName", locData[0].nameofpractice.ToLower().Contains("dr") ? locData[0].nameofpractice : locData[0].nameofpractice);
                    body = body.Replace("#address", locData[0].address);
                    //body = body.Replace("#Address", locData[0].address);
                    body = body.Replace("#Address", locData[0].address + "<br/>" + locData[0].city + ", " + locData[0].state + " " + locData[0].zipcode);
                    body = body.Replace("#location", locData[0].location);
                    body = body.Replace("#Nameofpractice", locData[0].nameofpractice.ToLower().Contains("dr") ? locData[0].nameofpractice : locData[0].nameofpractice);
                    body = body.Replace("#Phone", locData[0].telephone);
                }
                //Preop printing







                if (postData != null)
                {
                    var CC = string.IsNullOrEmpty(postData.txtHistoryPresentillness) ? "" : postData.txtHistoryPresentillness;

                    CC = CC.Replace("#dos", data);


                    body = body.Replace("#CC", CC);
                    //body = body.Replace("#CC", string.IsNullOrEmpty(postData.txtHistoryPresentillness) ? "" : this.removePtag(postData.txtHistoryPresentillness));

                    body = body.Replace("#PhysicalExamination", string.IsNullOrEmpty(postData.txtPhysicalExamination) ? "" : this.removePtag(postData.txtPhysicalExamination));
                    body = body.Replace("#TREATMENT", string.IsNullOrEmpty(postData.txtExaminedResult) ? "" : this.removePtag(postData.txtExaminedResult));
                    if (patientData.compensation == "WC")
                    {
                        var dd = string.IsNullOrEmpty(page1Data.dd) ? "" : page1Data.dd;
                        var workstatus = string.IsNullOrEmpty(page1Data.work_status) ? "" : page1Data.work_status;
                        var impart = "<p><b>IMPAIRMENT RATING: </b>" + dd + "" + workstatus + "</p>";

                        body = body.Replace("#impart", impart);

                    }
                    else
                    {
                        body = body.Replace("#impart", "");
                    }

                    ViewBag.ieId = patientData.id;
                    ViewBag.fuId = fuid;
                    ViewBag.locId = patientData.location_id;


                    string signName = "";
                    int signUserId = 0;

                    int? providorId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId);

                    if (providerData.Id != null)
                    {
                        signUserId = providerData.Id.Value;
                    }
                    else if (providorId != null)
                    {
                        signUserId = providorId.Value;
                    }

                    if (signUserId > 0)
                    {
                        tbl_users _user = new tbl_users()
                        {
                            Id = signUserId
                        };
                        var userData = _userService.GetOne(_user);
                        signName = userData.signature;

                        if (!string.IsNullOrEmpty(signName))
                        {
                            string signatureUrl = $"/Uploads/Sign/" + cmpid + "/" + signName;
                            //string signatureUrl = "https://paintrax.com/newversionlive/Uploads/Sign/" + cmpid + "/" + signName;
                            string base64Image = ImageToBase64(Environment.WebRootPath + signatureUrl);
                            body = body.Replace("#Sign", $" <img src='data:image/jpg;base64,{base64Image}' alt='My Image' />");
                            // body = body.Replace("#Sign", $"<img crossorigin='anonymous|use-credentials' src='{signatureUrl}' alt='Patient Signature' />");
                        }
                        else
                            body = body.Replace("#Sign", "");

                        body = body.Replace("#Physician", providerData.providername);

                        body = body.Replace("#ProviderName", providerData.providername.ToUpper());
                        body = body.Replace("#AssProviderName", providerData.assistant_providername);
                    }
                    else
                        body = body.Replace("#Sign", "");
                }
                body = body.Replace("#br", "<br/>");
                ViewBag.content = body;
            }
            catch (Exception ex)
            {
                SaveLog(ex, "PostOpPrint");
            }
            return View();
        }
        #endregion
        #region PrintFU

        public IActionResult FUPrint(int ieid, int fuid)
        {

            //try
            //{
            string plan = "";
            ViewBag.url = HttpContext.Request.GetEncodedUrl();

            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
            string body = string.Empty;
            //using streamreader for reading my htmltemplate   

            var fuData = _patientFuservices.GetOne(fuid);
            var fed = _patientFuservices.GetFirstIEDateOne(fuData.patientIE_ID.Value);

            var templateData = _printService.GetTemplate(cmpid, fuData.type);
            var gender = "";

            body = templateData.content;

            var patientData = _ieService.GetOnebyPatientId(ieid);

            if (patientData != null)
            {
                gender = Common.GetMrMrsFromSex(patientData.gender);
                body = body.Replace("#wcbgroup", patientData.primary_wcb_group);
                body = body.Replace("#patientname", gender + " " + patientData.fname + " " + patientData.mname + " " + patientData.lname);
                body = body.Replace("#fn", patientData.fname);
                body = body.Replace("#ln", patientData.lname);
                body = body.Replace("#dob", Common.commonDate(patientData.dob));
                body = body.Replace("#doi", Common.commonDate(patientData.doa));
                body = body.Replace("#dos", Common.commonDate(fuData.doe, HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));
                body = body.Replace("#location", patientData.location);
                body = body.Replace("#age", patientData.age == null ? "0" : patientData.age.Value.ToString());
                body = body.Replace("#gender", gender);
                body = body.Replace("#pp", string.IsNullOrEmpty(fuData.procedure_performed) ? "" : fuData.procedure_performed.Replace("^", "<br/>"));
                body = body.Replace("#sex", Common.GetGenderFromSex(patientData.gender));


                // body = body.Replace("#CT", System.Enum.GetName(typeof(CaseType), Convert.ToInt32(patientData.compensation)));
                body = body.Replace("#CT", patientData.compensation);
                body = body.Replace("#fed", fed == null ? "" : Common.commonDate(fed.Value, HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));
            }

            //header printing

            var locData = _locService.GetAll(" and id=" + patientData.location_id);

            if (locData != null && locData.Count > 0)
            {
                if (locData[0].nameofpractice != null)
                    body = body.Replace("#drFLName", locData[0].drfname + " " + locData[0].drlname);
                else
                    body = body.Replace("#drFLName", "");
                body = body.Replace("#drName", locData[0].nameofpractice == null ? "" : (locData[0].nameofpractice.ToLower().Contains("dr") ? locData[0].nameofpractice : locData[0].nameofpractice));
                body = body.Replace("#address", locData[0].address);
                //body = body.Replace("#Address", locData[0].address);
                body = body.Replace("#Address", locData[0].address + "<br/>" + locData[0].city + ", " + locData[0].state + " " + locData[0].zipcode);
                body = body.Replace("#loc", locData[0].location);
                body = body.Replace("#Location", locData[0].location);
                body = body.Replace("#Nameofpractice", locData[0].nameofpractice == null ? "" : (locData[0].nameofpractice.ToLower().Contains("dr") ? locData[0].nameofpractice : locData[0].nameofpractice));
                body = body.Replace("#Phone", locData[0].telephone);

                if (locData[0].drfname != null)
                    body = body.Replace("#NameofDr", locData[0].drfname + " " + locData[0].drlname);
                else
                    body = body.Replace("#NameofDr", "");


                string drAddress = "";
                if (!string.IsNullOrEmpty(locData[0].address))
                    drAddress = locData[0].address;

                if (!string.IsNullOrEmpty(locData[0].city))
                    drAddress = drAddress + "<br/>" + locData[0].city + " " + locData[0].state + " " + locData[0].zipcode;

                body = body.Replace("#DrAddress", drAddress);

            }

            //ADL printing

            var page2Data = _fuPage2services.GetOne(fuid);

            if (page2Data != null)
            {
                body = body.Replace("#ADL", this.removePtag(page2Data.aod));
                body = body.Replace("#Ros", this.removePtag(page2Data.ros));
            }
            else
            {
                body.Replace("#ADL", "");
                body = body.Replace("#Ros", "");
            }


            var page1DataIE = _ieService.GetOnePage1(ieid);

            //POC printing

            var pocData = this.getPOCPrint(fuid);

            //CC printing

            var page1Data = _fuPage1services.GetOne(fuid);
            if (page1Data != null)
            {
                body = body.Replace("#Reason", string.IsNullOrEmpty(page1Data.appt_reason) ? "" : this.removePtag(page1Data.appt_reason));


                string cc = "";
                string pe = "";

                cc = string.IsNullOrEmpty(page1Data.cc) ? "" : this.removePtag(page1Data.cc);
                pe = string.IsNullOrEmpty(page1Data.pe) ? "" : page1Data.pe;

                if (pocData != null)
                {
                    cc = string.IsNullOrEmpty(pocData.strCCDesc) ? cc : cc + "<br/><br/>" + pocData.strCCDesc;
                    pe = string.IsNullOrEmpty(pocData.strPEDesc) ? pe : pe + "<br/><br/>" + pocData.strPEDesc;
                }


                body = body.Replace("#CC", HtmlCleaner.ClearPE(cc));
                body = body.Replace("#PE", HtmlCleaner.ClearPE(pe));
                var history = string.IsNullOrEmpty(page1Data.history) ? "" : page1Data.history;
                var hstryPresentIllness = string.IsNullOrEmpty(page1Data.history) ? "" : "<p><b>HISTORY OF PRESENT ILLNESS: </b>" + page1Data.history.Replace("<p>", "<span>").Replace("</p>", "</span>") + "</p>";

                history = history.Replace("#dos", Common.commonDate(patientData.doe, HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));
                history = history.Replace("#doi", Common.commonDate(patientData.doa, HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));
                //  history.Replace("#patientname", gender + " " + patientData.fname + " " + patientData.mname + " " + patientData.lname);
                history = history.Replace("#patientname", gender + " " + patientData.lname + " " + patientData.fname + " " + patientData.lname);
                history = history.Replace("#accidenttype", patientData.accidentType);
                history = history.Replace("#accidenttype", patientData.accidentType);
                body = body.Replace("#history", history);

                body = body.Replace("#age", patientData.age == null ? "0" : patientData.age.Value.ToString());
                body = body.Replace("#DD", string.IsNullOrEmpty(page1Data.dd) ? "" : page1Data.dd);
                body = body.Replace("#WorkStatus", string.IsNullOrEmpty(page1Data.work_status) ? "" : page1Data.work_status);
                body = body.Replace("#POPlan", string.IsNullOrEmpty(page1Data.plan) ? "" : page1Data.plan);
                body = body.Replace("#fn", patientData.fname);
                body = body.Replace("#ln", patientData.lname);
                body = body.Replace("#gender", Common.GetMrMrsFromSex(patientData.gender));
                body = body.Replace("#sex", Common.GetGenderFromSex(patientData.gender));
                body = body.Replace("#accountno", patientData.account_no);

                if (patientData.compensation == "WC")
                {
                    var dd = string.IsNullOrEmpty(page1Data.dd) ? "" : page1Data.dd;
                    var workstatus = string.IsNullOrEmpty(page1Data.work_status) ? "" : page1Data.work_status;
                    var impart = "<p><b>IMPAIRMENT RATING: </b>" + dd + "" + workstatus + "</p>";

                    body = body.Replace("#impart", impart);

                }
                else
                {
                    body = body.Replace("#impart", "");
                }
                plan = string.IsNullOrEmpty(page1Data.plan) ? "" : page1Data.plan;
                string bodypart = "";

                if (!string.IsNullOrEmpty(page1Data.bodypart))
                    bodypart = Common.ReplceCommowithAnd(page1Data.bodypart.ToLower());

                body = body.Replace("#PC", string.IsNullOrEmpty(bodypart) ? "" : Common.FirstCharToUpper(bodypart) + " pain.");
                body = body.Replace("#bodypart", string.IsNullOrEmpty(bodypart) ? "" : bodypart.Replace(",", ", "));
                hstryPresentIllness = hstryPresentIllness.Replace("#bodypart", string.IsNullOrEmpty(bodypart) ? "" : bodypart.Replace(",", ", "));
                body = body.Replace("#hstryPresentIllness", hstryPresentIllness);

                string assessment = "";
                if (!string.IsNullOrEmpty(page1Data.assessment))
                {
                    if (!string.IsNullOrEmpty(bodypart))
                        assessment = page1Data.assessment.Replace("#PC", Common.FirstCharToUpper(bodypart) + " pain.");
                    else
                        assessment = page1Data.assessment.Replace("#PC", "");
                    assessment = assessment.Replace("#accidenttype", fuData.accident_type);
                }


                if (pocData != null)
                {
                    assessment = string.IsNullOrEmpty(pocData.strADesc) ? assessment : assessment + "<br/><br/>" + pocData.strADesc;
                }

                body = body.Replace("#doi", Common.commonDate(patientData.doa, HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));

                body = body.Replace("#PastMedicalHistory", this.removePtag(page1Data.pmh));
                body = body.Replace("#PastSurgicalHistory", this.removePtag(page1Data.psh));
                body = body.Replace("#SocialHistory", this.removePtag(page1Data.social_history));
                body = body.Replace("#Allergies", this.removePtag(page1Data.allergies));
                body = body.Replace("#FamilyHistory", this.removePtag(page1Data.family_history));
                body = body.Replace("#Vital", this.removePtag(page1Data.vital));
                body = body.Replace("#Diagnoses", this.removePtag(assessment));
                body = body.Replace("#Occupation", this.removePtag(page1Data.occupation));
                body = body.Replace("#PastMedications", this.removePtag(page1Data.medication));
                body = body.Replace("#DD", this.removePtag(page1Data.dd));
                body = body.Replace("#WorkStatus", this.removePtag(page1Data.work_status));
                body = body.Replace("#IR", this.removePtag(page1Data.impairment_rating));
                body = body.Replace("#age", patientData.age == null ? "0" : patientData.age.Value.ToString());
                body = body.Replace("#doi", Common.commonDate(patientData.doa, HttpContext.Session.GetString(SessionKeys.SessionDateFormat)));

            }
            else
            {
                body = body.Replace("#WorkStatus", "");
                body = body.Replace("#DD", "");
                body = body.Replace("#PastMedications", "");
                body = body.Replace("#Occupation", "");
                body = body.Replace("#Diagnoses", "");
                body = body.Replace("#Vital", "");
                body = body.Replace("#Allergies", "");
                body = body.Replace("#FamilyHistory", "");
                body = body.Replace("#PastSurgicalHistory", "");
                body = body.Replace("#SocialHistory", "");
                body = body.Replace("#PastMedicalHistory", "");
                body = body.Replace("#IR", "");
            }

            //last line 
            string lastline = "";
            if (page1Data != null)
                lastline = "It is my opinion that the injuries and symptoms " + gender + " " + patientData.fname + " " + patientData.mname + " " + patientData.lname + " sustained to " + page1Data.bodypart + " are causally related to the incident that occurred on " + Common.commonDate(patientData.doa) + " as described by the patient.";
            body = body.Replace("#lastline", lastline);

            //NE printing

            var pageNEData = _fuNEservices.GetOne(fuid);
            if (pageNEData != null)
            {

                body = body.Replace("#Sen_Exm", this.removePtag(pageNEData.sensory));
                body = body.Replace("#SE", this.removePtag(pageNEData.sensory));

                body = body.Replace("#MMST", this.removePtag(pageNEData.manual_muscle_strength_testing));

                body = body.Replace("#NE", this.removePtag(pageNEData.neurological_exam));

                body = body.Replace("#DTR", this.removePtag(pageNEData.other_content));

            }
            else
            {
                body = body.Replace("#DTR", "");
                body = body.Replace("#SE", "");
                body = body.Replace("#NE", "");
                body = body.Replace("#MMST", "");
                body = body.Replace("#Sen_Exm", "");
            }

            //NE printing

            var page3Data = _fuPage3services.GetOne(fuid);
            if (page3Data != null)
            {
                body = body.Replace("#Gait", this.removePtag(page3Data.gait));
                body = body.Replace("#Care", this.removePtag(page3Data.care));
                body = body.Replace("#Procedures", this.removePtag(page3Data.universal));
                body = body.Replace("#Goal", this.removePtag(page3Data.goal));
            }
            else
            {
                body = body.Replace("#Goal", "");
                body = body.Replace("#Procedures", "");
                body = body.Replace("#Care", "");
                body = body.Replace("#Gait", "");
            }

            //Treatment printing

            var pageOtherData = _fuOtherService.GetOne(fuid);
            if (pageOtherData != null)
            {

                body = body.Replace("#Treatment", this.removePtag(pageOtherData.treatment_details));
                body = body.Replace("#heshe", Common.GethesheFromSex(patientData.gender));
                body = body.Replace("#hisher", Common.GethisherFromSex(patientData.gender));
                body = body.Replace("#note1", this.removePtag(pageOtherData.note1));
                body = body.Replace("#note2", this.removePtag(pageOtherData.note2));
                body = body.Replace("#note3", this.removePtag(pageOtherData.note3));


                string fup_duration = "";

                if (!string.IsNullOrEmpty(pageOtherData.followup_duration))
                    fup_duration = pageOtherData.followup_duration;
                else if (pageOtherData.followup_date.HasValue)
                    fup_duration = Common.commonDate(pageOtherData.followup_date);


                body = body.Replace("#FollowUp", this.removePtag(fup_duration));
                body = body.Replace("#fup", this.removePtag(fup_duration));
            }
            else
            {
                body = body.Replace("#FollowUp", "");
                body = body.Replace("#Treatment", "");
            }

            //POC printing

            var dataPOC = this.getPOCPrint(fuid);

            if (string.IsNullOrEmpty(plan))
                body = body.Replace("#Plan", this.removePtag(dataPOC.strPoc));
            else
                body = body.Replace("#Plan", this.removePtag(plan));

            //var dataPOC = this.getPOC(ieid);



            //body = body.Replace("#Plan", this.removePtag(dataPOC.strPoc));

            body = body.Replace("#ReflexExam", "");
            string injectionHtml = dataPOC.strInjectionDesc;
            //string injectionHtml = "<h2>Injection Test</h2>";

            //string SessionDiffPage = HttpContext.Session.GetString(SessionKeys.SessionPageBreak);
            string SessionDiffDoc = HttpContext.Session.GetString(SessionKeys.SessionInjectionAsSeparateBlock);

            if (SessionDiffDoc != "true")
            {

                if (HttpContext.Session.GetString(SessionKeys.SessionPageBreak) == "true")
                {
                    if (!string.IsNullOrEmpty(injectionHtml))
                    {
                        // Create HTML with a page break before the injection section
                        string pageBreakHtml = "<div style='page-break-before: always;'>";
                        pageBreakHtml += injectionHtml;
                        pageBreakHtml += "</div>";

                        body = body.Replace("#injection", pageBreakHtml);
                    }
                    else
                    {
                        body = body.Replace("#injection", "");
                    }
                }
                else
                {
                    body = body.Replace("#injection", injectionHtml);
                }
            }
            else
            {
                body = body.Replace("#injection", "");
            }


            body = body.Replace("#ReflexExam", "");


            var strDiagnostic = this.getDiagnostic(fuid);


            if (string.IsNullOrEmpty(strDiagnostic))
            {
                if (HttpContext.Session.GetString(SessionKeys.SessionIsDaignosis) == "true")
                {
                    strDiagnostic = HttpContext.Session.GetString(SessionKeys.SessionDaignosisNotFoundStatment);
                }
            }
            else
            {
                if (HttpContext.Session.GetString(SessionKeys.SessionIsDaignosis) == "true")
                {
                    strDiagnostic = strDiagnostic + "<br/><br/>" + HttpContext.Session.GetString(SessionKeys.SessionDaignosisFoundStatment); ;
                }
            }


            body = body.Replace("#Diagnostic", this.removePtag(strDiagnostic));


            var data = _fuPage3services.GetOne(fuid);

            if (data != null)
            {
                body = body.Replace("#PrescribedMedications", data.discharge_medications);
            }
            else
            {
                body = body.Replace("#PrescribedMedications", "");
            }

            body = body.Replace("#LastNote", "");
            body = body.Replace("#sex", Common.GetGenderFromSex(patientData.gender));
            body = body.Replace("#gender", Common.GetMrMrsFromSex(patientData.gender));

            ViewBag.ieId = patientData.id;
            ViewBag.fuId = fuid;
            ViewBag.locId = patientData.location_id;


            string signName = "";
            int signUserId = 0;

            int? providorId = HttpContext.Session.GetInt32(SessionKeys.SessionSelectedProviderId);

            if (fuData.provider_id != null)
            {
                signUserId = fuData.provider_id.Value;
            }
            else if (providorId != null)
            {
                signUserId = providorId.Value;
            }
            body = HtmlCleaner.ClearHTML(body);
            ViewBag.content = body;
            if (signUserId > 0)
            {
                tbl_users _user = new tbl_users()
                {
                    Id = signUserId
                };
                var userData = _userService.GetOne(_user);
                signName = userData.signature;

                if (!string.IsNullOrEmpty(signName))
                {
                    string signatureUrl = $"/Uploads/Sign/" + cmpid + "/" + signName;
                    //string signatureUrl = "https://paintrax.com/newversionlive/Uploads/Sign/" + cmpid + "/" + signName;
                    string base64Image = ImageToBase64(Environment.WebRootPath + signatureUrl);
                    body = body.Replace("#Sign", $" <img src='data:image/jpg;base64,{base64Image}' alt='My Image' />");
                    // body = body.Replace("#Sign", $"<img crossorigin='anonymous|use-credentials' src='{signatureUrl}' alt='Patient Signature' />");
                }
                else
                    body = body.Replace("#Sign", "");

                body = body.Replace("#ProviderName", userData.providername);
                body = body.Replace("#AssProviderName", userData.assistant_providername);
                ViewBag.ProviderName = userData.providername;
            }
            else
                body = body.Replace("#Sign", "");

            body = body.Replace("<ol>", "<ol style='font-family: 'Times New Roman', Times, serif;'>");
            if (HttpContext.Session.GetString(SessionKeys.SessionShowTableBorder) == "false")
            {
                body = body.Replace("<table>", "<table style='width:100%;border:none;'>");
                body = body.Replace("<tr>", "<tr style='border:none;'>");
                body = body.Replace("<td>", "<td style='border:none;'>");
            }
            else
            {
                body = body.Replace("<table>", "<table style='width:100%;'>");
            }
            //body = HtmlCleaner.CleanHtmlContent(body);

            if (SessionDiffDoc == "true")
            {
                body += "<br><br><!--Diff Doc-->";
                body += injectionHtml;
            }

            // string updatedHtml = HtmlCleaner.ClearHTML(body);

            ViewBag.content = body;

            //}
            //catch (Exception ex)
            //{
            //    SaveLog(ex, "FUPrint");
            //}
            return View();
        }

        #region Image to Base64
        public string ImageToBase64(string imagePath)
        {
            string base64String = string.Empty;
            if (System.IO.File.Exists(imagePath))
            {
                byte[] imageBytes = System.IO.File.ReadAllBytes(imagePath);
                base64String = Convert.ToBase64String(imageBytes);
                // Use base64String as needed
            }

            return base64String;
        }
        #endregion

        [HttpPost]
        public IActionResult DownloadPDF(String htmlContent, string ieId)
        {
            //if (!string.IsNullOrEmpty(htmlContent))
            //{
            //    IronPdf.HtmlToPdf Renderer = new IronPdf.HtmlToPdf();
            //    Renderer.RenderHtmlAsPdf(htmlContent).SaveAs("html-string.pdf");
            //    return Ok();
            //}
            //else
            //{
            //    var stream = new FileStream(Path.Combine(_hostingEnvironment.ContentRootPath, "html-string.pdf"), FileMode.Open);
            //    return new FileStreamResult(stream, "application/pdf");
            //}

            return View();

        }

        [HttpPost]
        public IActionResult DownloadWord(string htmlContent, int ieId, int fuId, string provName = "")
        {
            htmlContent = htmlContent.Replace("<p>&nbsp;</p>", "");



            //  string htmlContent = "<p>This is a <strong>sample</strong> HTML content.</p>";
            string filePath = "", docName = "", patientName = "", injDocName = "", dos = "", dob = "";
            string[] splitContent;
            string injHtmlContent = "";
            if (SessionDiffDoc == "true")
            {

                splitContent = htmlContent.Split(new string[] { "<!--Diff Doc-->" }, StringSplitOptions.None);
                htmlContent = splitContent[0];
                if (splitContent.Length > 1)
                    injHtmlContent = splitContent[1];

            }
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

                    // Convert HTML to OpenXML and add to the body
                    //HtmlConverter converter = new HtmlConverter(mainPart);
                    //var generatedBody = converter.Parse(htmlContent);
                    //body.Append(generatedBody);


                    var header = new Header(new Paragraph(new Run(new Text("Header Test"))));
                    HeaderReference headerReference = new HeaderReference() { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) };
                    //var footer = new Footer(new Paragraph(new Run(new Text("Page"), new SimpleField() { Instruction = "PAGE" })));
                    //FooterReference footerReference = new FooterReference() { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) };

                    headerPart.Header = header;

                    mainPart.Document.Body.Append(new SectionProperties(headerReference));

                }
                string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

                var patientData = _ieService.GetOnebyPatientId(ieId);
                var fuData = _patientFuservices.GetOne(fuId);
                if (fuData.type == "Preop H and P")
                {
                    if (patientData.doa == null)
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_PreOP_" + Common.commonDate(fuData.doe).Replace("/", "") + ".docx";
                    }
                    else if (patientData.account_no != null)
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_PreOP_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + patientData.account_no + "_" + Common.commonDate(patientData.doa).Replace("/", "") + ".docx";
                    }
                    else
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_PreOP_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + Common.commonDate(patientData.doa).Replace("/", "") + ".docx";
                    }
                }
                else if (fuData.type == "Postop FU")
                {
                    if (patientData.doa == null)
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_POP_" + Common.commonDate(fuData.doe).Replace("/", "") + ".docx";
                    }
                    else if (patientData.account_no != null)
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_POP_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + patientData.account_no + "_" + Common.commonDate(patientData.doa).Replace("/", "") + ".docx";
                    }
                    else
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_POP_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + Common.commonDate(patientData.doa).Replace("/", "") + ".docx";
                    }
                }
                else
                {
                    if (patientData.doa == null)
                    {
                        if (patientData.account_no != null)
                            docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_FU_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + patientData.account_no + ".docx";
                        else
                            docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_FU_" + Common.commonDate(fuData.doe).Replace("/", "") + ".docx";
                    }

                    else if (patientData.account_no != null)
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_FU_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + patientData.account_no + "_" + Common.commonDate(patientData.doa).Replace("/", "") + ".docx";
                    }

                    else
                    {
                        docName = patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_FU_" + Common.commonDate(fuData.doe).Replace("/", "") + "_" + Common.commonDate(patientData.doa).Replace("/", "") + ".docx";
                    }
                }

                patientName = patientData.lname.Trim() + ", " + patientData.fname.Trim();

                dos = fuData.doe == null ? "" : fuData.doe.Value.ToShortDateString();
                dob = patientData.dob == null ? "" : patientData.dob.Value.ToShortDateString();

                string subPath = "Report/" + cmpid; // Your code goes here

                bool exists = System.IO.Directory.Exists(subPath);

                if (!exists)
                    System.IO.Directory.CreateDirectory(subPath);

                filePath = subPath + "/" + docName;

                // Save the memory stream to a file
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                {

                    memStream.WriteTo(fileStream);
                }
                injDocName = subPath + "/" + patientData.lname.Trim() + ", " + patientData.fname.Trim() + "_FU_" + Common.commonDate(patientData.doe).Replace("/", "") + "_injection.docx";
                System.IO.File.Delete(injDocName);
                injHtmlContent = Regex.Replace(injHtmlContent, @"</ol>", string.Empty, RegexOptions.IgnoreCase);
                injHtmlContent = injHtmlContent.Trim();

                if (injHtmlContent != "")
                {

                    MemoryStream injMemStream = ConvertHtmlToWord(injHtmlContent);
                    // injDocName = subPath + "/"+ patientData.lname + "," + patientData.fname + "_IE_" + Common.commonDate(patientData.doe).Replace("/", "") + "_injection.docx";
                    using (FileStream fileStream = new FileStream(injDocName, FileMode.Create))
                    {

                        injMemStream.WriteTo(fileStream);
                    }
                }
            }


            return Json(new { filePath = filePath, fileName = docName, patientName = patientName, dos = dos, dob = dob, injFileName = injDocName, provName = provName });
        }

        public MemoryStream ConvertHtmlToWord(string htmlContent)
        {
            MemoryStream memoryStream = new MemoryStream();
            using (WordprocessingDocument doc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                // Create the MainDocumentPart
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
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

                // Save the changes
                mainPart.Document.Save();
            }
            return memoryStream;
        }

        [HttpGet]
        public virtual ActionResult DownloadFile(string filePath, string fileName, int locId = 0, string patientName = "", string signatureUrl = "", string dos = "", string dob = "", string injFileName = "")
        {
            string cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();

            var dt = _locService.GetAll(" and cmp_id=" + cmpid + " and id=" + locId);

            if (dt.Count > 0)
            {
                if (!string.IsNullOrEmpty(dt[0].header_template))
                {
                    string filepathFrom = Path.Combine(Environment.WebRootPath, "Uploads/HeaderTemplate") + "//" + dt[0].header_template;


                    string filepathTo = filePath;
                    AddHeaderFromTo(filepathFrom, filepathTo, patientName, dos);
                    if (DoesFooterExist(filepathFrom))
                        AddFooterFromTo(filepathFrom, filepathTo, patientName, dos, dob);

                }
                else
                {
                    try
                    {
                        string filepathFrom = Path.Combine(Environment.WebRootPath, "Uploads/HeaderTemplate") + "//" + HttpContext.Session.GetString(SessionKeys.SessionHeaderTemplate); ;


                        string filepathTo = filePath;
                        AddHeaderFromTo(filepathFrom, filepathTo, patientName, dos);
                        if (DoesFooterExist(filepathFrom))
                            AddFooterFromTo(filepathFrom, filepathTo, patientName, dos, dob);
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            byte[] data = System.IO.File.ReadAllBytes(filePath);
            return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

        }

        //public static void AddHeaderFromTo(string filepathFrom, string filepathTo)
        //{
        //    // Replace header in target document with header of source document.
        //    using (WordprocessingDocument
        //        wdDoc = WordprocessingDocument.Open(filepathTo, true))
        //    {
        //        MainDocumentPart mainPart = wdDoc.MainDocumentPart;

        //        // Delete the existing header part.
        //        mainPart.DeleteParts(mainPart.HeaderParts);

        //        // Create a new header part.
        //        DocumentFormat.OpenXml.Packaging.HeaderPart headerPart =
        //    mainPart.AddNewPart<HeaderPart>();

        //        // Get Id of the headerPart.
        //        string rId = mainPart.GetIdOfPart(headerPart);

        //        // Feed target headerPart with source headerPart.
        //        using (WordprocessingDocument wdDocSource =
        //            WordprocessingDocument.Open(filepathFrom, true))
        //        {
        //            DocumentFormat.OpenXml.Packaging.HeaderPart firstHeader =
        //    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

        //            wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

        //            if (firstHeader != null)
        //            {
        //                headerPart.FeedData(firstHeader.GetStream());
        //            }
        //        }

        //        // Get SectionProperties and Replace HeaderReference with new Id.
        //        IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs =
        //    mainPart.Document.Body.Elements<SectionProperties>();
        //        foreach (var sectPr in sectPrs)
        //        {
        //            // Delete existing references to headers.
        //            sectPr.RemoveAllChildren<HeaderReference>();

        //            // Create the new header reference node.
        //            sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
        //        }
        //    }
        //}

        public void AddHeaderFromTo(string filepathFrom, string filepathTo, string patientName = "", string dos = "")
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
                }

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                var restheaderPart = mainPart.AddNewPart<HeaderPart>("Rest");
                restheaderPart.Header = CreateHeaderWithPageNumber(patientName, "");
                if (cmpid == 7 || cmpid == 13)
                {
                    if (!string.IsNullOrEmpty(dos))
                    {
                        string _dos = Common.commonDate(Convert.ToDateTime(dos), HttpContext.Session.GetString(SessionKeys.SessionDateFormat));
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
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Type = HeaderFooterValues.Default, Id = restId });
                }
            }
        }
        public Header CreateHeaderWithPageNumber(string text1, string text2)
        {
            if (text2 != "")
            {
                return new Header(
                    new Paragraph(
                        new Run(
                            new Text(text1) // First line
                        ),
                        new Break(), // Line break
                        new Run(
                            new Text(text2) // Second line
                        ),
                        new Break(), // Another line break if needed
                        new Run(
                            new Text("Page ") // Static "Page " text
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
                return new Header(
                   new Paragraph(
                       new Run(
                           new Text(text1) // First line
                       ),
                       new Break(), // Line break

                       new Run(
                           new Text("Page ") // Static "Page " text
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
        }

        public Footer CreateFooterWithPageNumber(string text1, string text2, string text3)
        {
            if (text2 != "")
            {
                return new Footer(
                    new Paragraph(
                        new Run(
                            new Text(text1) // First line
                        ),
                          new Run(new TabChar()), new Run(new TabChar()), new Run(new TabChar()),
                        new Run(
                            new Text(text2) // Second line
                        ),
                          new Run(new TabChar()), new Run(new TabChar()), new Run(new TabChar()),
                        new Run(
                            new Text(text3) // Second line
                        ),
                          new Run(new TabChar()), new Run(new TabChar()), new Run(new TabChar()),
                        new Run(
                            new Text("Page ") // Static "Page " text
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
                return new Footer(
                   new Paragraph(
                       new Run(
                           new Text(text1) // First line
                       ),
                          new Break(),
                       new Run(
                           new Text("Page ") // Static "Page " text
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
        }

        #endregion

        #region Private Method
        private int calculateAge(DateTime bday, DateTime? dos)
        {
            DateTime today = dos == null ? DateTime.Today : dos.Value;

            int age = today.Year - bday.Year;

            if (today.Month < bday.Month ||
       ((today.Month == bday.Month) && (today.Day < bday.Day)))
            {
                age--;  //birthday in current year not yet reached, we are 1 year younger ;)
                        //+ no birthday for 29.2. guys ... sorry, just wrong date for birth
            }

            return age;
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

        private string removePtag(string content)
        {
            if (!string.IsNullOrEmpty(content))
            {
                Regex rgx = new Regex("<p>|</p>");
                string res = rgx.Replace(content, "", 2);
                return res;
            }
            else
                return "";
        }

        private string getTreatment(string val)
        {
            string returnStr = "";
            if (!string.IsNullOrEmpty(val))
            {
                string[] str = val.Split('`');



                for (int i = 0; i < str.Length; i++)
                {

                    if (!string.IsNullOrEmpty(str[i]) && str[i].Substring(0, 1) != "@")
                    {
                        returnStr += "<br/>" + str[i].TrimStart('@');
                    }
                }
            }
            return returnStr;
        }

        private pocDetails getPOC(int PatientFU_ID)
        {
            DataTable dsPOC = _pocService.GetPOCFU(PatientFU_ID);

            string strPoc = "", pocDesc = "", ccdesc = "", pedesc = "", adesc = "";
            string inject_desc = "", pageBreakHtml = "";
            if (dsPOC != null && dsPOC.Rows.Count > 0)
            {

                for (int i = 0; i < dsPOC.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(dsPOC.Rows[i]["Heading"].ToString()))
                    {

                        //if (i != dsPOC.Tables[0].Rows.Count - 1)
                        //    strPoc = strPoc + "<b style='text-transform:uppercase'>" + dsPOC.Tables[0].Rows[i]["Heading"].ToString().TrimEnd(':') + ": </b>" + dsPOC.Tables[0].Rows[i]["PDesc"].ToString() + "<br/><br/>";
                        //else

                        //if (dsPOC.Tables[0].Rows[i]["Executed"] != DBNull.Value)
                        //{
                        //    this.generatePNReport(bodypart, dsPOC.Tables[0].Rows[i]["MCODE"].ToString(), ViewState["name"].ToString(),
                        //        ViewState["dob"].ToString(), ViewState["location"].ToString(), "", dsPOC.Tables[0].Rows[i]["Level"].ToString(), "", "", ViewState["doe"].ToString(), PatientIE_ID);
                        //}

                        string heading = dsPOC.Rows[i]["Heading"].ToString();
                        pocDesc = dsPOC.Rows[i]["PDesc"].ToString();
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());




                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());


                        adesc = adesc + dsPOC.Rows[i]["ADesc"] == null ? "" : dsPOC.Rows[i]["ADesc"].ToString().Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());


                        if (heading.ToLower().Contains("(side)"))
                        {
                            heading = heading.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString().ToUpper());
                            heading = heading.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString().ToUpper());

                        }

                        if (pocDesc.ToLower().Contains("(side)"))
                        {
                            pocDesc = pocDesc.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            pocDesc = pocDesc.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        }

                        if (heading.ToLower().Contains("(levels)"))
                        {
                            heading = heading.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            heading = heading.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (pocDesc.ToLower().Contains("(levels)"))
                        {
                            pocDesc = pocDesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            pocDesc = pocDesc.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (heading.ToLower().Contains("(level)"))
                        {
                            heading = heading.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            heading = heading.Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());

                        }

                        if (pocDesc.ToLower().Contains("(level)"))
                        {
                            pocDesc = pocDesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            pocDesc = pocDesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (dsPOC.Rows[i]["pn"].ToString() == "1" && dsPOC.Rows[i]["Executed"] != DBNull.Value)
                        {
                            if (!string.IsNullOrEmpty(dsPOC.Rows[i]["injection_description"].ToString()))
                            {
                                inject_desc = (dsPOC.Rows[i]["injection_description"].ToString());
                                inject_desc = inject_desc.Replace("#Side", dsPOC.Rows[i]["Sides"].ToString());
                                inject_desc = inject_desc.Replace("#Muscle", dsPOC.Rows[i]["Muscle"].ToString().TrimEnd('~').ToString().Replace("~", ", "));
                                inject_desc = inject_desc.Replace("#Level", dsPOC.Rows[i]["Level"].ToString().TrimEnd('~').ToString().Replace("~", ", "));
                                inject_desc = inject_desc.Replace("#Medication", dsPOC.Rows[i]["Medication"].ToString().TrimEnd('~').ToString().Replace("~", ", "));

                                inject_desc = "<div style='page-break-before: always;'>" + inject_desc + "</div>";
                                pageBreakHtml = pageBreakHtml + inject_desc;

                            }
                        }
                        // strPoc = strPoc + "<li><b style='text-transform:uppercase'>" + heading.TrimEnd(':') + ": </b>" + pocDesc + "</li>";
                        strPoc = strPoc + "<li><b>" + heading + " </b>" + pocDesc + "</li>";
                    }
                }
            }

            pocDetails pocDetails = new pocDetails()
            {
                strInjectionDesc = pageBreakHtml,
                strPoc = strPoc != "" ? "<ol>" + strPoc + "</ol>" : "",
                strCCDesc = ccdesc,
                strPEDesc = pedesc,
                strADesc = adesc
            };

            return pocDetails;
        }

        private pocDetails getPOCPrint(int PatientFU_ID)
        {
            DataTable dsPOC = _pocService.GetPOCFUPrint(PatientFU_ID);



            string strPoc = "";
            string inject_desc = "", ccdesc = "", pedesc = "", adesc = "", pocDesc = "", pageBreakHtml = "";
            if (dsPOC != null && dsPOC.Rows.Count > 0)
            {
                int cnt = 0;
                for (int i = 0; i < dsPOC.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(dsPOC.Rows[i]["Heading"].ToString()))
                    {

                        string heading = dsPOC.Rows[i]["Heading"].ToString();
                        pocDesc = dsPOC.Rows[i]["PDesc"].ToString();
                        string _ccdesc = dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString();
                        if (_ccdesc != "")
                        {
                            _ccdesc = _ccdesc.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            _ccdesc = _ccdesc.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                            _ccdesc = _ccdesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            _ccdesc = _ccdesc.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                            _ccdesc = _ccdesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            _ccdesc = _ccdesc.Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());

                            ccdesc = ccdesc + _ccdesc;
                        }


                        string _pedesc = dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString();
                        if (_pedesc != "")
                        {
                            _pedesc = _pedesc.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            _pedesc = _pedesc.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                            _pedesc = _pedesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            _pedesc = _pedesc.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                            _pedesc = _pedesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            _pedesc = _pedesc.Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());

                            pedesc = pedesc + _pedesc;
                        }

                        string _adesc = dsPOC.Rows[i]["ADesc"] == null ? "" : dsPOC.Rows[i]["ADesc"].ToString();

                        if (_adesc != null)
                        {
                            _adesc = _adesc.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            _adesc = _adesc.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                            _adesc = _adesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            _adesc = _adesc.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                            _adesc = _adesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            _adesc = _adesc.Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());

                            adesc = adesc + _adesc;
                        }

                        if (heading.ToLower().Contains("(side)"))
                        {
                            heading = heading.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            heading = heading.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        }
                        if (pocDesc.ToLower().Contains("(side)"))
                        {
                            pocDesc = pocDesc.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            pocDesc = pocDesc.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        }
                        if (pocDesc.ToLower().Contains("(levels)"))
                        {
                            pocDesc = pocDesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            pocDesc = pocDesc.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        }
                        if (pocDesc.ToLower().Contains("(level)"))
                        {
                            pocDesc = pocDesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            pocDesc = pocDesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        }


                        if (heading.ToLower().Contains("(levels)"))
                        {
                            heading = heading.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            heading = heading.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (heading.ToLower().Contains("(level)"))
                        {
                            heading = heading.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            heading = heading.Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (dsPOC.Rows[i]["pn"].ToString() == "1" && dsPOC.Rows[i]["Executed"] != DBNull.Value)
                        {
                            if (!string.IsNullOrEmpty(dsPOC.Rows[i]["injection_description"].ToString()))
                            {
                                inject_desc = (dsPOC.Rows[i]["injection_description"].ToString());
                                inject_desc = inject_desc.Replace("#Side", dsPOC.Rows[i]["Sides"].ToString());
                                inject_desc = inject_desc.Replace("#Muscle", dsPOC.Rows[i]["Muscle"].ToString().TrimEnd('~').ToString().Replace("~", ", "));
                                inject_desc = inject_desc.Replace("#Level", dsPOC.Rows[i]["Level"].ToString().TrimEnd('~').ToString().Replace("~", ", "));
                                inject_desc = inject_desc.Replace("#Medication", dsPOC.Rows[i]["Medication"].ToString().TrimEnd('~').ToString().Replace("~", ", "));

                                if (cnt > 0)
                                    inject_desc = "<div style='page-break-before: always;'>" + inject_desc + "</div>";

                                pageBreakHtml = pageBreakHtml + inject_desc;
                                cnt++;

                            }
                        }
                        //strPoc = strPoc + "<li><b style='text-transform:uppercase'>" + heading.TrimEnd(':') + ": </b>" + pocDesc + "</li>";
                        strPoc = strPoc + "<li><b >" + heading + " </b>" + pocDesc + "</li>";
                    }
                }
            }

            pocDetails pocDetails = new pocDetails()
            {
                strInjectionDesc = pageBreakHtml,
                strPoc = strPoc != "" ? "<ol>" + strPoc + "</ol>" : "",
                strCCDesc = ccdesc,
                strPEDesc = pedesc,
                strADesc = adesc
            };

            return pocDetails;
        }

        private pocDetails UpdatePOCPlan(int PatientFU_ID)
        {
            DataTable dsPOC = _pocService.GetPOCFUPrint(PatientFU_ID);



            string strPoc = "", pocDesc = "", ccdesc = "", pedesc = "", adesc = "";
            string inject_desc = "";
            if (dsPOC != null && dsPOC.Rows.Count > 0)
            {

                for (int i = 0; i < dsPOC.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(dsPOC.Rows[i]["Heading"].ToString()))
                    {


                        string heading = dsPOC.Rows[i]["Heading"].ToString();
                        pocDesc = dsPOC.Rows[i]["PDesc"].ToString();
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        ccdesc = ccdesc + dsPOC.Rows[i]["CCDesc"] == null ? "" : dsPOC.Rows[i]["CCDesc"].ToString().Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());




                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        pedesc = pedesc + dsPOC.Rows[i]["PEDesc"] == null ? "" : dsPOC.Rows[i]["PEDesc"].ToString().Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());


                        adesc = adesc + dsPOC.Rows[i]["ADesc"] == null ? "" : dsPOC.Rows[i]["ADesc"].ToString().Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        adesc = adesc + dsPOC.Rows[i]["adesc"] == null ? "" : dsPOC.Rows[i]["adesc"].ToString().Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());





                        if (heading.ToLower().Contains("(side)"))
                        {
                            heading = heading.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString().ToUpper());
                            heading = heading.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString().ToUpper());

                        }

                        if (pocDesc.ToLower().Contains("(side)"))
                        {
                            pocDesc = pocDesc.Replace("(SIDE)", dsPOC.Rows[i]["Sides"].ToString());
                            pocDesc = pocDesc.Replace("(side)", dsPOC.Rows[i]["Sides"].ToString());
                        }

                        if (heading.ToLower().Contains("(levels)"))
                        {
                            heading = heading.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            heading = heading.Replace("(LEVELS)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (pocDesc.ToLower().Contains("(levels)"))
                        {
                            pocDesc = pocDesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                            pocDesc = pocDesc.Replace("(levels)", dsPOC.Rows[i]["Level"].ToString());
                        }

                        if (heading.ToLower().Contains("(level)"))
                        {
                            heading = heading.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            heading = heading.Replace("(LEVEL)", dsPOC.Rows[i]["Level"].ToString());

                        }

                        if (pocDesc.ToLower().Contains("(level)"))
                        {
                            pocDesc = pocDesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                            pocDesc = pocDesc.Replace("(level)", dsPOC.Rows[i]["Level"].ToString());
                        }





                        if (dsPOC.Rows[i]["pn"].ToString() == "1" && dsPOC.Rows[i]["Executed"] != DBNull.Value)
                        {
                            if (!string.IsNullOrEmpty(dsPOC.Rows[i]["injection_description"].ToString()))
                            {
                                inject_desc = inject_desc + "<br/>" + (dsPOC.Rows[i]["injection_description"].ToString());
                                inject_desc = inject_desc.Replace("#Side", dsPOC.Rows[i]["Sides"].ToString());
                                inject_desc = inject_desc.Replace("#Muscle", dsPOC.Rows[i]["Muscle"].ToString().TrimEnd('~').ToString().Replace("~", ", "));
                            }
                        }
                        //strPoc = strPoc + "<li><b style='text-transform:uppercase'>" + heading + " </b>" + pocDesc + "</li>";
                        strPoc = strPoc + "<li><b >" + heading + " </b>" + pocDesc + "</li>";
                    }
                }
            }

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            if (cmpid.Value == 18 || cmpid.Value == 15 || cmpid.Value == 10 || cmpid.Value == 2)
                strPoc = strPoc.Replace("<li>", "").Replace("</li>", "<br/>");

            pocDetails pocDetails = new pocDetails()
            {
                strInjectionDesc = inject_desc,
                strPoc = strPoc != "" ? ((cmpid.Value == 18 || cmpid.Value == 15 || cmpid.Value == 10 || cmpid.Value == 2) ? strPoc : "<ol>" + strPoc + "</ol>") : "",
                strCCDesc = ccdesc,
                strPEDesc = pedesc,
                strADesc = adesc
            };

            return pocDetails;
        }
        private string getPOCDate(int fuid, int ieId)
        {
            string potion = null;
            string iinew = string.Empty;
            string test = "";
            var postData = _fuPostService.GetOne(fuid);
            if (postData != null)
            {
                if (postData.chkLeftKnee == true)
                {
                    potion = "Left";
                    iinew = "knee";
                }
                else if (postData.chkRightKnee == true)
                {
                    potion = "Right";
                    iinew = "knee";
                }
                else if (postData.chkLeftShoulder == true)
                {
                    potion = "Left";
                    iinew = "shoulder";
                }
                else
                {
                    potion = "Right";
                    iinew = "shoulder";
                }
            }

            int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            //var x = _pocService.GetAllProceduresFU(iinew.Trim(), patientFUId, potion, cmpid.Value); //commented by moulick as all poc required so specific visit wise poc cancled. 
            var x = _pocService.GetAllProcedures(iinew.Trim(), ieId, potion, cmpid.Value);
            if (x != null)
            {
                if (x.Rows.Count > 0)
                {
                    foreach (DataRow row in x.Rows)
                    {

                        foreach (DataColumn column in x.Columns)
                        {
                            if (column.ColumnName == "Executed")
                            {
                                if (row[17] != DBNull.Value)
                                {
                                    test = row[17] != DBNull.Value ? Convert.ToDateTime(row[17]).ToString("MM/dd/yyyy").Replace('-', '/') : string.Empty;
                                }

                            }
                        }
                    }
                }
            }
            //var x = _pocService.GetAllProcedures(iinew.Trim(), patientIEId, potion);

            //if (x != null)
            //{
            //    if (x.Rows.Count > 0)
            //    {
            //        ViewBag.executeddate = Convert.ToDateTime(row[17]).ToString("MM/dd/yyyy").Replace('-', '/');

            //    }
            //}

            return test;
        }
        private string getDiagnosticie(int id, tbl_pre pre)
        {
            var data = _ieService.GetOnePage3(id);

            string strDaignosis = "", stradddaigno = "", strCommaValue = "";
            bool isnormal = true;

            if (data != null)
            {

                if (data.diagcervialbulge_date != null)
                {
                    strDaignosis = data.diagcervialbulge_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagcervialbulge_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagcervialbulge_study));

                        strDaignosis = strDaignosis + " " + study;
                    }

                    //if (!string.IsNullOrEmpty(data.diagcervialbulge_text))
                    //{

                    if (data.diagcervialbulge_comma != null)
                        strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagcervialbulge_comma));
                    strDaignosis = strDaignosis + " of the cervical spine " + strCommaValue + " " + data.diagcervialbulge_text;


                    stradddaigno = stradddaigno + "Cervical " + (string.IsNullOrEmpty(data.diagcervialbulge_text) ? "" : data.diagcervialbulge_text.Replace("reveals", "").TrimEnd('.')) + ".<br/>";
                    if (!string.IsNullOrEmpty(data.diagcervialbulge_text))
                        isnormal = false;
                    //}

                    if (!string.IsNullOrEmpty(data.diagcervialbulge_hnp1))
                    {
                        strDaignosis = strDaignosis + " HNP at " + data.diagcervialbulge_hnp1.TrimEnd('.') + ".";
                        stradddaigno = stradddaigno + "Cervical herniated nucleus pulposis at " + data.diagcervialbulge_hnp1.TrimEnd('.') + ".<br/>";
                        isnormal = false;
                    }

                    if (!string.IsNullOrEmpty(data.diagcervialbulge_hnp2))
                    {
                        strDaignosis = strDaignosis + data.diagcervialbulge_hnp2.TrimEnd('.') + ".";
                        if (!string.IsNullOrEmpty(data.diagcervialbulge_hnp2))
                        {
                            stradddaigno = stradddaigno + data.diagcervialbulge_hnp2.TrimEnd('.') + ".<br/>";
                        }

                        isnormal = false;
                    }

                    if (isnormal)
                    {
                        strDaignosis = strDaignosis + " normal study. ";
                    }
                }

                if (data.diagthoracicbulge_date != null)
                {
                    isnormal = true;
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagthoracicbulge_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagthoracicbulge_study));
                        strDaignosis = strDaignosis + " " + study;
                    }

                    //if (!string.IsNullOrEmpty(data.diagthoracicbulge_text))
                    //{
                    strCommaValue = "";
                    if (data.diagthoracicbulge_comma != null)
                        strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagthoracicbulge_comma));
                    strDaignosis = strDaignosis + " of the thoracic spine " + strCommaValue + " " + data.diagthoracicbulge_text;

                    stradddaigno = stradddaigno + "Thoracic " + (string.IsNullOrEmpty(data.diagthoracicbulge_text) ? "" : data.diagthoracicbulge_text.ToString().Replace("reveals", "").TrimEnd('.')) + ".<br/>";
                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_text))
                        isnormal = false;
                    //}

                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_hnp1))
                    {
                        strDaignosis = strDaignosis + " HNP at " + data.diagthoracicbulge_hnp1.TrimEnd('.') + ". ";
                        stradddaigno = stradddaigno + "Thoracic herniated nucleus pulposis at " + data.diagthoracicbulge_hnp1.TrimEnd('.') + ".<br/>";
                        isnormal = false;
                    }

                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_hnp2))
                    {
                        strDaignosis = strDaignosis + data.diagthoracicbulge_hnp2.TrimEnd('.') + ". ";
                        if (!string.IsNullOrEmpty(data.diagthoracicbulge_hnp2))
                        {
                            stradddaigno = stradddaigno + data.diagthoracicbulge_hnp2.TrimEnd('.') + ".<br/>";
                        }

                        isnormal = false;
                    }

                    if (isnormal)
                    {
                        strDaignosis = strDaignosis + " normal study. ";
                    }
                }

                if (data.diaglumberbulge_date != null)
                {
                    isnormal = true;
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diaglumberbulge_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diaglumberbulge_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diaglumberbulge_study));
                        strDaignosis = strDaignosis + " " + study;
                    }

                    //if (!string.IsNullOrEmpty(data.diaglumberbulge_text))
                    //{
                    strCommaValue = "";
                    if (data.diaglumberbulge_comma != null)
                        strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diaglumberbulge_comma));
                    strDaignosis = strDaignosis + " of the lumbar spine " + strCommaValue + " " + data.diaglumberbulge_text;

                    stradddaigno = stradddaigno + "Lumbar " + (string.IsNullOrEmpty(data.diaglumberbulge_text) ? "" : data.diaglumberbulge_text.ToString().Replace("reveals", "").TrimEnd('.')) + ".<br/>";
                    if (!string.IsNullOrEmpty(data.diaglumberbulge_text))
                        isnormal = false;
                    //}

                    if (!string.IsNullOrEmpty(data.diaglumberbulge_hnp1))
                    {
                        strDaignosis = strDaignosis + " HNP at " + data.diaglumberbulge_hnp1.TrimEnd('.') + ". ";
                        stradddaigno = stradddaigno + "Lumbar herniated nucleus pulposis at " + data.diaglumberbulge_hnp1.TrimEnd('.') + ".<br/>";
                        isnormal = false;
                    }

                    if (!string.IsNullOrEmpty(data.diaglumberbulge_hnp2))
                    {
                        strDaignosis = strDaignosis + data.diaglumberbulge_hnp2.TrimEnd('.') + ". ";
                        if (!string.IsNullOrEmpty(data.diaglumberbulge_hnp2))
                        {
                            stradddaigno = stradddaigno + data.diaglumberbulge_hnp2.TrimEnd('.') + ".<br/>";
                        }

                        isnormal = false;
                    }

                    if (isnormal)
                    {
                        strDaignosis = strDaignosis + " normal study. ";
                    }
                }
                if (data.diagrightshoulder_date != null)
                {

                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagrightshoulder_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagrightshoulder_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagrightshoulder_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagrightshoulder_text))
                    {
                        strCommaValue = "";
                        if (data.diagrightshoulder_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagrightshoulder_comma));

                        strDaignosis = strDaignosis + " of the right shoulder" + strCommaValue + " " + data.diagrightshoulder_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the right shoulder is normal study. ";
                    }

                }

                if (data.diagleftshoulder_date != null)
                {

                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagleftshoulder_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagleftshoulder_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagleftshoulder_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagleftshoulder_text))
                    {
                        strCommaValue = "";
                        if (data.diagleftshoulder_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagleftshoulder_comma));

                        strDaignosis = strDaignosis + " of the left shoulder" + strCommaValue + " " + data.diagleftshoulder_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + "  of the left shoulder is normal study. ";
                    }

                }



                if (data.diagrightknee_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagrightknee_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagrightknee_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagrightknee_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagrightknee_text))
                    {
                        strCommaValue = "";
                        if (data.diagrightknee_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagrightknee_comma));

                        strDaignosis = strDaignosis + " of the right knee" + strCommaValue + " " + data.diagrightknee_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the right knee is normal study. ";
                    }
                }
                if (data.diagleftknee_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagleftknee_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagleftknee_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagleftknee_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagleftknee_text))
                    {
                        strCommaValue = "";
                        if (data.diagleftknee_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagleftknee_comma));

                        strDaignosis = strDaignosis + " of the left knee" + strCommaValue + " " + data.diagleftknee_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the left knee is normal study. ";
                    }
                }



                if (data.other1_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other1_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other1_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other1_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other1_text))
                    {
                        strCommaValue = "";
                        if (data.other1_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other1_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other1_text.TrimEnd('.') + ". ";
                    }
                }

                if (data.other2_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other2_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other2_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other2_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other2_text))
                    {
                        strCommaValue = "";
                        if (data.other2_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other2_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other2_text.TrimEnd('.') + ". ";
                    }
                }

                if (data.other3_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other3_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other3_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other3_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other3_text))
                    {
                        strCommaValue = "";
                        if (data.other3_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other3_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other3_text.TrimEnd('.') + ". ";
                    }
                }

                if (data.other4_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other4_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other4_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other4_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other4_text))
                    {
                        strCommaValue = "";
                        if (data.other4_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other4_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other4_text.TrimEnd('.') + ". ";
                    }
                }

                if (data.other5_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other5_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other5_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other5_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other5_text))
                    {
                        strCommaValue = "";
                        if (data.other5_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other5_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other5_text.TrimEnd('.') + ". ";
                    }
                }

                if (data.other6_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other6_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other6_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other6_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other6_text))
                    {
                        strCommaValue = "";
                        if (data.other6_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other6_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other6_text.TrimEnd('.') + ". ";
                    }
                }

                if (data.other7_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other7_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other7_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other7_study));
                        strDaignosis = strDaignosis + " " + study + " ";
                    }

                    if (!string.IsNullOrEmpty(data.other7_text))
                    {
                        strCommaValue = "";
                        if (data.other7_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other7_comma));

                        strDaignosis = strDaignosis + " " + strCommaValue + " " + data.other7_text.TrimEnd('.') + ". ";
                    }
                }

            }

            return strDaignosis;

        }
        private string getDiagnostic(int id)
        {
            var data = _fuPage3services.GetOne(id);

            string strDaignosis = "", stradddaigno = "", strCommaValue = "";
            bool isnormal = true;

            if (data != null)
            {

                if (data.diagcervialbulge_date != null)
                {
                    strDaignosis = data.diagcervialbulge_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagcervialbulge_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagcervialbulge_study));

                        strDaignosis = strDaignosis + " " + study;
                    }

                    //if (!string.IsNullOrEmpty(data.diagcervialbulge_text))
                    //{

                    if (data.diagcervialbulge_comma != null)
                        strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagcervialbulge_comma));
                    strDaignosis = strDaignosis + " of the cervical spine " + strCommaValue + " " + data.diagcervialbulge_text;


                    stradddaigno = stradddaigno + "Cervical " + (data.diagcervialbulge_text == null ? "" : data.diagcervialbulge_text.Replace("reveals", "").TrimEnd('.')) + ".<br/>";
                    if (!string.IsNullOrEmpty(data.diagcervialbulge_text))
                        isnormal = false;
                    //}

                    if (!string.IsNullOrEmpty(data.diagcervialbulge_hnp1))
                    {
                        strDaignosis = strDaignosis + " HNP at " + data.diagcervialbulge_hnp1.TrimEnd('.') + ".";
                        stradddaigno = stradddaigno + " Cervical herniated nucleus pulposis at " + data.diagcervialbulge_hnp1.TrimEnd('.') + ".<br/>";
                        isnormal = false;
                    }

                    if (!string.IsNullOrEmpty(data.diagcervialbulge_hnp2))
                    {
                        strDaignosis = strDaignosis + data.diagcervialbulge_hnp2.TrimEnd('.') + ".";
                        if (!string.IsNullOrEmpty(data.diagcervialbulge_hnp2))
                        {
                            stradddaigno = stradddaigno + data.diagcervialbulge_hnp2.TrimEnd('.') + ".<br/>";
                        }

                        isnormal = false;
                    }

                    if (isnormal)
                    {
                        strDaignosis = strDaignosis + " normal study. ";
                    }
                }

                if (data.diagthoracicbulge_date != null)
                {
                    isnormal = true;
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagthoracicbulge_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagthoracicbulge_study));
                        strDaignosis = strDaignosis + " " + study;
                    }

                    //if (!string.IsNullOrEmpty(data.diagthoracicbulge_text))
                    //{
                    strCommaValue = "";
                    if (data.diagthoracicbulge_comma != null)
                        strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagthoracicbulge_comma));
                    strDaignosis = strDaignosis + " of the thoracic spine " + strCommaValue + " " + data.diagthoracicbulge_text;

                    stradddaigno = stradddaigno + "Thoracic " + (data.diagthoracicbulge_text == null ? "" : data.diagthoracicbulge_text.ToString().Replace("reveals", "").TrimEnd('.')) + ".<br/>";
                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_text))
                        isnormal = false;
                    //}

                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_hnp1))
                    {
                        strDaignosis = strDaignosis + " HNP at " + data.diagthoracicbulge_hnp1.TrimEnd('.') + ". ";
                        stradddaigno = stradddaigno + "Thoracic herniated nucleus pulposis at " + data.diagthoracicbulge_hnp1.TrimEnd('.') + ".<br/>";
                        isnormal = false;
                    }

                    if (!string.IsNullOrEmpty(data.diagthoracicbulge_hnp2))
                    {
                        strDaignosis = strDaignosis + data.diagthoracicbulge_hnp2.TrimEnd('.') + ". ";
                        if (!string.IsNullOrEmpty(data.diagthoracicbulge_hnp2))
                        {
                            stradddaigno = stradddaigno + data.diagthoracicbulge_hnp2.TrimEnd('.') + ".<br/>";
                        }

                        isnormal = false;
                    }

                    if (isnormal)
                    {
                        strDaignosis = strDaignosis + " normal study. ";
                    }
                }

                if (data.diaglumberbulge_date != null)
                {
                    isnormal = true;
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diaglumberbulge_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diaglumberbulge_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diaglumberbulge_study));
                        strDaignosis = strDaignosis + " " + study;
                    }

                    //if (!string.IsNullOrEmpty(data.diaglumberbulge_text))
                    //{
                    strCommaValue = "";
                    if (data.diaglumberbulge_comma != null)
                        strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diaglumberbulge_comma));
                    strDaignosis = strDaignosis + " of the lumbar spine " + strCommaValue + " " + data.diaglumberbulge_text;

                    stradddaigno = stradddaigno + " Lumbar " + (data.diaglumberbulge_text == null ? "" : data.diaglumberbulge_text.ToString().Replace("reveals", "").TrimEnd('.')) + ".<br/>";
                    if (!string.IsNullOrEmpty(data.diaglumberbulge_text))
                        isnormal = false;
                    //}

                    if (!string.IsNullOrEmpty(data.diaglumberbulge_hnp1))
                    {
                        strDaignosis = strDaignosis + " HNP at " + data.diaglumberbulge_hnp1.TrimEnd('.') + ". ";
                        stradddaigno = stradddaigno + " Lumbar herniated nucleus pulposis at " + data.diaglumberbulge_hnp1.TrimEnd('.') + ".<br/>";
                        isnormal = false;
                    }

                    if (!string.IsNullOrEmpty(data.diaglumberbulge_hnp2))
                    {
                        strDaignosis = strDaignosis + data.diaglumberbulge_hnp2.TrimEnd('.') + ". ";
                        if (!string.IsNullOrEmpty(data.diaglumberbulge_hnp2))
                        {
                            stradddaigno = stradddaigno + data.diaglumberbulge_hnp2.TrimEnd('.') + ".<br/>";
                        }

                        isnormal = false;
                    }

                    if (isnormal)
                    {
                        strDaignosis = strDaignosis + " normal study. ";
                    }
                }
                if (data.diagrightshoulder_date != null)
                {

                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagrightshoulder_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagrightshoulder_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagrightshoulder_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagrightshoulder_text))
                    {
                        strCommaValue = "";
                        if (data.diagrightshoulder_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagrightshoulder_comma));

                        strDaignosis = strDaignosis + " of the right shoulder " + strCommaValue + " " + data.diagrightshoulder_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the right shoulder is normal study. ";
                    }

                }
                if (data.diagleftshoulder_date != null)
                {

                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagleftshoulder_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagleftshoulder_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagleftshoulder_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagleftshoulder_text))
                    {
                        strCommaValue = "";
                        if (data.diagleftshoulder_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagleftshoulder_comma));

                        strDaignosis = strDaignosis + " of the left shoulder " + strCommaValue + " " + data.diagleftshoulder_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the left shoulder is normal study. ";
                    }

                }


                if (data.diagrightknee_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagrightknee_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagrightknee_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagrightknee_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagrightknee_text))
                    {
                        strCommaValue = "";
                        if (data.diagrightknee_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagrightknee_comma));

                        strDaignosis = strDaignosis + " of the right knee " + strCommaValue + " " + data.diagrightknee_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the right knee is normal study. ";
                    }
                }

                if (data.diagleftknee_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.diagleftknee_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.diagleftknee_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study1>(data.diagleftknee_study));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.diagleftknee_text))
                    {
                        strCommaValue = "";
                        if (data.diagleftknee_comma != null)
                            strCommaValue = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.diagleftknee_comma));

                        strDaignosis = strDaignosis + " of the left knee " + strCommaValue + " " + data.diagleftknee_text.TrimEnd('.') + ". " + "<br/>";
                    }
                    else
                    {
                        strDaignosis = strDaignosis + " of the left knee is normal study. ";
                    }
                }



                if (data.other1_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other1_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other1_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other1_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other1_text))
                    {
                        strCommaValue = "";
                        if (data.other1_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other1_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other1_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

                if (data.other2_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other2_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other2_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other2_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other2_text))
                    {
                        strCommaValue = "";
                        if (data.other2_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other2_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other2_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

                if (data.other3_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other3_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other3_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other3_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other3_text))
                    {
                        strCommaValue = "";
                        if (data.other3_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other3_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other3_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

                if (data.other4_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other4_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other4_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other4_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other4_text))
                    {
                        strCommaValue = "";
                        if (data.other4_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other4_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other4_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

                if (data.other5_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other5_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other5_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other5_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other5_text))
                    {
                        strCommaValue = "";
                        if (data.other5_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other5_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other5_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

                if (data.other6_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other6_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other6_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other6_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other6_text))
                    {
                        strCommaValue = "";
                        if (data.other6_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other6_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other6_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

                if (data.other7_date != null)
                {
                    strDaignosis = (!string.IsNullOrEmpty(strDaignosis) ? (strDaignosis + "<br/>") : "") + data.other7_date.Value.ToString("MM/dd/yyyy") + " - ";

                    if (!string.IsNullOrEmpty(data.other7_study))
                    {
                        var study = EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.Study2>(data.other7_study.Replace(" ", "").Replace("/", "_")));
                        strDaignosis = strDaignosis + study;
                    }

                    if (!string.IsNullOrEmpty(data.other7_text))
                    {
                        strCommaValue = "";
                        if (data.other7_comma != null)
                            strCommaValue = " " + EnumHelper.GetDisplayName(System.Enum.Parse<EnumHelper.StudyComma>(data.other7_comma));

                        strDaignosis = strDaignosis + strCommaValue + " " + data.other7_text.TrimEnd('.') + ". " + "<br/>";
                    }
                }

            }

            return strDaignosis;

        }
        public bool DoesFooterExist(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) // Open as read-only
            {
                return wordDoc.MainDocumentPart.FooterParts.Any();
            }
        }
        public void AddFooterFromTo(string filepathFrom, string filepathTo, string patientName = "", string dos = "", string dob = "")
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument
                wdDoc = WordprocessingDocument.Open(filepathTo, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                // Delete the existing footer part.
                mainPart.DeleteParts(mainPart.FooterParts);

                // Create a new footer part.
                DocumentFormat.OpenXml.Packaging.FooterPart footerPart =
            mainPart.AddNewPart<FooterPart>();
                // Get Id of the footerPart.
                string rId = mainPart.GetIdOfPart(footerPart);

                // Feed target headerPart with source headerPart.
                using (WordprocessingDocument wdDocSource =
                    WordprocessingDocument.Open(filepathFrom, true))
                {
                    DocumentFormat.OpenXml.Packaging.FooterPart firstFooter =
            wdDocSource.MainDocumentPart.FooterParts.FirstOrDefault();



                    if (firstFooter != null)
                    {
                        footerPart.FeedData(firstFooter.GetStream());
                    }
                    Dictionary<string, string> imageRelMapping = new Dictionary<string, string>();

                    foreach (var imagePart in firstFooter.ImageParts)
                    {
                        // Add a new image part to the target footer
                        ImagePart newImagePart = footerPart.AddImagePart(imagePart.ContentType);

                        // Copy the image data
                        using (Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            newImagePart.FeedData(imageStream);
                        }

                        // Map the old relationship ID to the new image part ID
                        string oldRelId = firstFooter.GetIdOfPart(imagePart);
                        string newRelId = footerPart.GetIdOfPart(newImagePart);
                        imageRelMapping[oldRelId] = newRelId;
                    }



                    // Update relationships in header XML
                    UpdateFooterXml(footerPart, imageRelMapping);
                }

                int? cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

                var restfooterPart = mainPart.AddNewPart<FooterPart>("RestFooter");
                // restfooterPart.Footer = CreateHeaderWithPageNumber(patientName, "");
                if (cmpid == 7 || cmpid == 13)
                {
                    if (!string.IsNullOrEmpty(dos))
                    {
                        string _dos = Common.commonDate(Convert.ToDateTime(dos), HttpContext.Session.GetString(SessionKeys.SessionDateFormat));
                        restfooterPart.Footer = CreateFooterWithPageNumber(patientName, _dos, "");
                    }
                }
                else if (cmpid == 15)
                {
                    if (!string.IsNullOrEmpty(dos))
                    {
                        string _dos = Common.commonDate(Convert.ToDateTime(dos), HttpContext.Session.GetString(SessionKeys.SessionDateFormat));
                        string _dob = Common.commonDate(Convert.ToDateTime(dob), HttpContext.Session.GetString(SessionKeys.SessionDateFormat));
                        restfooterPart.Footer = CreateFooterWithPageNumber(patientName, _dos, _dob);
                    }
                }
                else
                {

                    restfooterPart.Footer = CreateFooterWithPageNumber(patientName, "", "");
                }


                //  restheaderPart.Header = new Header(new Paragraph("Purav\nSandip"));
                string restId = mainPart.GetIdOfPart(restfooterPart);
                // Get SectionProperties and Replace HeaderReference with new Id.
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs = mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to footers.
                    sectPr.RemoveAllChildren<FooterReference>();
                    sectPr.Append(new TitlePage());
                    // Create the new footer reference node.
                    sectPr.PrependChild<FooterReference>(new FooterReference() { Type = HeaderFooterValues.First, Id = rId });
                    sectPr.PrependChild<FooterReference>(new FooterReference() { Type = HeaderFooterValues.Default, Id = restId });
                }


            }
        }
        // Method to update header XML to reference new image relationships
        private static void UpdateFooterXml(FooterPart footerPart, Dictionary<string, string> imageRelMapping)
        {
            string headerXml;

            // Read the existing header XML
            using (StreamReader reader = new StreamReader(footerPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                headerXml = reader.ReadToEnd();
            }

            // Replace old relationship IDs with new ones
            foreach (var kvp in imageRelMapping)
            {
                headerXml = headerXml.Replace($"r:id=\"{kvp.Key}\"", $"r:id=\"{kvp.Value}\"");
            }

            // Write the updated XML back to the header part
            using (MemoryStream memStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(memStream))
                {
                    writer.Write(headerXml);
                    writer.Flush();
                    memStream.Position = 0;
                    footerPart.FeedData(memStream);
                }
            }
        }

        [HttpGet]
        public JsonResult SaveProcedurePerformed(int eId, string pp)
        {
            _patientFuservices.UpdateProcedurePerformed(eId.ToString(), pp);

            return Json(1);
        }
        #endregion
    }
}
