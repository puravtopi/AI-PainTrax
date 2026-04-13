using Microsoft.AspNetCore.Mvc.Rendering;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace PainTrax.Web.Helper
{
    public class Common
    {
        private readonly DesinationServices _designationservices = new DesinationServices();
        private readonly GroupsService _groupservices = new GroupsService();
        private readonly LocationsService _locationservices = new LocationsService();
        private readonly POCConfigService _pocconfigservices = new POCConfigService();
        private readonly CaseTypeService _caseTypeService = new CaseTypeService();
        private readonly UserService _userservices = new UserService();
        private readonly CommonService _commonService = new CommonService();
        private readonly AccidentTypeService _accidentTypeService = new AccidentTypeService();
        private readonly StateService _stateService = new StateService();
        private readonly ReferringPhysicianService _physicianService = new ReferringPhysicianService();
        private readonly VisitTypeService _visitTypeService = new VisitTypeService();
        private readonly ProSXServices _proSXServices = new ProSXServices();

        public List<SelectListItem> GetDesignation(int cmp_id)
        {

            string cnd = " and cmp_id=" + cmp_id;

            var dgdata = _designationservices.GetAll(cnd);
            var dglist = new List<SelectListItem>();
            dglist.Add(new SelectListItem
            {
                Text = "--Select Designation--",
                Value = ""
            });

            foreach (var item in dgdata)
            {
                dglist.Add(new SelectListItem
                {
                    Text = item.title,
                    Value = item.id.ToString()
                });
            }


            return dglist;
        }

        public List<SelectListItem> GetGroups(int cmp_id)
        {
            string cnd = " and cmp_id=" + cmp_id;
            var gpdata = _groupservices.GetAll(cnd);
            var gplist = new List<SelectListItem>();
            gplist.Add(new SelectListItem
            {
                Text = "--Select Group--",
                Value = ""
            });

            foreach (var item in gpdata)
            {
                gplist.Add(new SelectListItem
                {
                    Text = item.Title,
                    Value = item.Id.ToString()
                });
            }

            return gplist;
        }

        public List<SelectListItem> GetLocations(int cmp_id)
        {
            string cnd = " and cmp_id=" + cmp_id+ " and isactive=1 order by location";
            var data = _locationservices.GetAll(cnd);
            var list = new List<SelectListItem>();


            list.Add(new SelectListItem
            {
                Text = "--Select Location--",
                Value = "0"
            });


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.location,
                    Value = item.id.ToString()
                });
            }

            return list;
        }


        //public List<SelectListItem> GetProSXReportDate(int cmp_id)
        //{

        //    var list = new List<SelectListItem>();
        //    var data = _proSXServices.GetProSXReportDate(cmp_id.ToString());

        //    list.Add(new SelectListItem
        //    {
        //        Text = "--Select Date--",
        //        Value = ""
        //    });


        //    foreach (var item in data)
        //    {
        //        list.Add(new SelectListItem
        //        {
        //            Text = item,
        //            Value = item
        //        });
        //    }

        //    return list;
        //}

        public List<SelectListItem> GetCaseType(int cmp_id)
        {
            string cnd = " and cmp_id=" + cmp_id;
            var data = _caseTypeService.GetAll(cnd);
            var list = new List<SelectListItem>();


            list.Add(new SelectListItem
            {
                Text = "--Select CaseType--",
                Value = ""
            });


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.casetype,
                    Value = item.casetype
                });
            }

            return list;
        }

        public List<SelectListItem> GetAccidenttype(int cmp_id)
        {
            string cnd = " and cmp_id=" + cmp_id;
            var data = _accidentTypeService.GetAll(cnd);
            var list = new List<SelectListItem>();


            list.Add(new SelectListItem
            {
                Text = "--Select Accident Type--",
                Value = ""
            });


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.accidenttype,
                    Value = item.accidenttype
                });
            }

            return list;
        }

        public List<SelectListItem> GetUsers(int cmp_id)
        {
            string cnd = " and cmp_id=" + cmp_id;
            var data = _userservices.GetAll(cnd);
            var list = new List<SelectListItem>();

            list.Add(new SelectListItem
            {
                Text = "--Select User--",
                Value = ""
            });

            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.fname + ' ' + item.lname,
                    Value = item.Id.ToString()
                });
            }

            return list;
        }
        public List<SelectListItem> GetState(int cmp_id)
        {
            // string cnd = " and cmp_id=" + cmp_id;
            string cnd = " ";
            var data = _stateService.GetAll(cnd);
            var list = new List<SelectListItem>();


            list.Add(new SelectListItem
            {
                Text = "--Select State--",
                Value = ""
            });


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.state_name + "-" + item.fullname,
                    Value = item.state_name
                });
            }

            return list;
        }
        public List<SelectListItem> GetPhysician(int cmp_id)
        {
            string cnd = " and cmp_id=" + cmp_id;
            var data = _physicianService.GetAll(cnd);
            var list = new List<SelectListItem>();


            list.Add(new SelectListItem
            {
                Text = "--Select Physician--",
                Value = ""
            });


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.physicianname,
                    Value = item.Id.ToString(),
                });
            }

            return list;
        }

        public List<CheckBoxItem> GetLocationsCheckBoxList(int cmp_id, string loc_ids = "")
        {
            string cnd = " and cmp_id=" + cmp_id;
            var data = _locationservices.GetAll(cnd);
            var list = new List<CheckBoxItem>();


            foreach (var item in data)
            {
                list.Add(new CheckBoxItem
                {
                    Item = item.location,
                    Id = item.id.Value,
                    IsChecked = loc_ids.Contains(item.id.ToString())
                });
            }

            return list;
        }
        public List<tbl_pocconfig> GetpocconfigColumnList(int cmp_id, string _ids = "")

        {

            string cnd = "";// " and cmp_id=" + cmp_id;

            var data = _pocconfigservices.GetAll(cnd);
            return data;
        }
        public List<CheckBoxItem> GetpocconfigCheckBoxList(int cmp_id, string _ids = "")

        {

            string cnd = " and cmp_id=" + cmp_id;

            var data = _pocconfigservices.GetAll(cnd);

            var list = new List<CheckBoxItem>();

            string input = "Sex,Name,Account_No,Case,DOE,DOB,DOA,MCODE,Side,Level,Phone,Location,Insurance,ClaimNo,MC,Allergies,Request,Scheduled,Executed,Note,SC_Name,Policy,Attorney,Vaccinated,Surgeon,Assistant";





            string[] entries = input.Split(',');

            int i = 0;

            foreach (var item in entries)

            {

                bool ischecked = false;

                foreach (var item1 in data)

                {

                    if (item1.columns == item.TrimEnd())

                    { ischecked = true; }





                }

                list.Add(new CheckBoxItem

                {

                    Item = item,

                    Id = i,

                    IsChecked = Convert.ToBoolean(ischecked)

                });

                i++;

            }





            //foreach (var item in data)

            //{

            //    list.Add(new CheckBoxItem

            //    {

            //        Item = item.columns,

            //        //    Id = item.columns,

            //        IsChecked = _ids.Contains(item.id.ToString())

            //    });

            //}



            return list;

        }
        public List<CheckBoxItem> GetpocconfigExportCheckBoxList(int cmp_id, string _ids = "")

        {

            string cnd = " and cmp_id=" + cmp_id;

            var data = _pocconfigservices.GetAllExport(cnd);

            var list = new List<CheckBoxItem>();

            string input = "Sex,Name,Account_No,Case,DOE,DOB,DOA,MCODE,Side,Level,Phone,Location,Insurance,ClaimNo,MC,Allergies,Request,Scheduled,Executed,Note,SC_Name,Policy,Attorney,Vaccinated,Surgeon,Assistant";





            string[] entries = input.Split(',');

            int i = 0;

            foreach (var item in entries)

            {

                bool ischecked = false;

                foreach (var item1 in data)

                {

                    if (item1.export_Columns == item.TrimEnd())

                    { ischecked = true; }





                }

                list.Add(new CheckBoxItem

                {

                    Item = item,

                    Id = i,

                    IsChecked = Convert.ToBoolean(ischecked)

                });

                i++;

            }         


            return list;

        }
        public List<CheckBoxItem> GetPagesCheckBoxList(string pages_ids = "")
        {

            var data = _commonService.GetPagesAll("");
            var list = new List<CheckBoxItem>();


            foreach (var item in data)
            {
                list.Add(new CheckBoxItem
                {
                    Item = item.page_name,
                    Id = item.id,
                    IsChecked = pages_ids.Contains(item.id.ToString())
                });
            }

            return list;
        }

        public List<CheckBoxItem> GetReportsCheckBoxList(string pages_ids = "")
        {

            var data = _commonService.GetReportsAll("");
            var list = new List<CheckBoxItem>();


            foreach (var item in data)
            {
                list.Add(new CheckBoxItem
                {
                    Item = item.report_name,
                    Id = item.id,
                    IsChecked = pages_ids.Contains(item.id.ToString())
                });
            }

            return list;
        }

        public List<CheckBoxItem> GetRolsCheckBoxList(string pages_ids = "")
        {

            var data = _commonService.GetRoleAll("");
            var list = new List<CheckBoxItem>();


            foreach (var item in data)
            {
                list.Add(new CheckBoxItem
                {
                    Item = item.name,
                    Id = item.id,
                    IsChecked = pages_ids.Contains(item.id.ToString())
                });
            }

            return list;
        }

        public List<SelectListItem> GetSOAPSettingList(string colNames)
        {

            var data = GetPage1Feilds();
            var list = new List<SelectListItem>();

            string[] colNameArray = colNames.Split(',');
            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Value = item.Value,
                    Text = item.Text,
                    Selected = colNameArray.Contains(item.Value, StringComparer.OrdinalIgnoreCase)
                });
            }

            return list;
        }

        public List<SelectListItem> GetPage2SettingList(string colNames)
        {

            var data = GetPage2Feilds();
            var list = new List<SelectListItem>();


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Value = item.Value,
                    Text = item.Text,
                    Selected = colNames.Contains(item.Value)
                });
            }

            return list;
        }
        public List<SelectListItem> GetPageNESettingList(string colNames)
        {

            var data = GetPageNEFeilds();
            var list = new List<SelectListItem>();


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Value = item.Value,
                    Text = item.Text,
                    Selected = colNames.Contains(item.Value)
                });
            }

            return list;
        }

        public List<SelectListItem> GetPageOtherSettingList(string colNames)
        {

            var data = GetPageOtherFeilds();
            var list = new List<SelectListItem>();


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Value = item.Value,
                    Text = item.Text,
                    Selected = colNames.Contains(item.Value)
                });
            }

            return list;
        }

        public List<SelectListItem> GetDaignosisFeilds(string colNames)
        {

            var data = GetDaignosisFeilds();
            var list = new List<SelectListItem>();


            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Value = item.Value,
                    Text = item.Text,
                    Selected = colNames.Contains(item.Value)
                });
            }

            return list;
        }

        private static string Encrypt(string clearText)
        {
            string encryptionKey = "PAINTRAX2023";
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(encryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return clearText;
        }

        private static string Decrypt(string cipherText)
        {
            string encryptionKey = "PAINTRAX2023";
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(encryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }

        public static DateTime? ConvertDate(string date)
        {
            if (string.IsNullOrEmpty(date) == false)
            {
                return Convert.ToDateTime(date.Replace("-", "/"));
            }
            else
                return null;
        }

        public static string commonDate(DateTime? date, string dateformat = "MM/dd/yyyy")
        {
            if (date != null)
                return Convert.ToDateTime(date).ToString(dateformat).Replace('-', '/');
            else
                return "";
        }

        public static string? GetMrMrsFromSex(string sex)
        {
            if (string.IsNullOrEmpty(sex) == false)
            {
                return sex == "1" ? "Mr." : "Ms.";
            }
            else
                return "";
        }
        public static string? GethesheFromSex(string sex)
        {
            if (string.IsNullOrEmpty(sex) == false)
            {
                return sex == "1" ? "he" : "she";
            }
            else
                return "";
        }
        public static string? GethisherFromSex(string sex)
        {
            if (string.IsNullOrEmpty(sex) == false)
            {
                return sex == "1" ? "his" : "her";
            }
            else
                return "";
        }

        public static string? GetGenderFromSex(string sex)
        {
            if (string.IsNullOrEmpty(sex) == false)
            {
                return sex == "1" ? "male" : "female";
            }
            else
                return "";
        }

        public static string? GetGenderFromMrMrs(string sex)
        {
            if (string.IsNullOrEmpty(sex) == false)
            {
                return sex == "Mr." ? "1" : "2";
            }
            else
                return "";
        }

        public static string FirstCharToUpper(string input)
        {
            if (String.IsNullOrEmpty(input))
                throw new ArgumentException("ARGH!");
            return input.First().ToString().ToUpper() + input.Substring(1);
        }

        public static string ReplceCommowithAnd(string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                var lastComma = input.LastIndexOf(',');
                if (lastComma != -1)
                    input = input.Remove(lastComma, 1).Insert(lastComma, " and ");

            }
            return input;
        }


        private List<SelectListItem> GetPage1Feilds()
        {

            List<SelectListItem> returnstatus = new List<SelectListItem>
            {


                new SelectListItem{ Text="History", Value = "history",Selected=false },
                new SelectListItem{ Text="Body Parts", Value = "bodypart",Selected=false },
                new SelectListItem{ Text="CE", Value = "cc",Selected=false },
                new SelectListItem{ Text="PE", Value = "pe",Selected=false },
                new SelectListItem{ Text="Assessment", Value = "assessment",Selected=false },
                new SelectListItem{ Text="Po Plan", Value = "plan",Selected=false },
                new SelectListItem{ Text="Occupation", Value = "occupation",Selected=false },
                new SelectListItem{ Text="Degree of Disability", Value = "dd",Selected=false },
                new SelectListItem{ Text="Work Status", Value = "work_status",Selected=false },
                new SelectListItem{ Text="PMH", Value = "pmh",Selected=false },
                new SelectListItem{ Text="PSH", Value = "psh",Selected=false },
                new SelectListItem{ Text="Allergies", Value = "allergies",Selected=false },
                new SelectListItem{ Text="Medication", Value = "medication",Selected=false },
                new SelectListItem{ Text="Family History", Value = "family_history",Selected=false },
                new SelectListItem{ Text="Social History", Value = "social_history",Selected=false },
                new SelectListItem{ Text="Impairment Rating", Value = "impairment_rating",Selected=false },
                new SelectListItem{ Text="Reason", Value = "appt_reason",Selected=false },
                new SelectListItem{ Text="Vital", Value = "vital",Selected=false },
                new SelectListItem{ Text="POC assessment", Value = "poc_assesment",Selected=false },
                new SelectListItem{ Text="Procedure Performed", Value = "procedure_performed",Selected=false }



                          };
            return returnstatus;

        }

        private List<SelectListItem> GetPage2Feilds()
        {

            List<SelectListItem> returnstatus = new List<SelectListItem>
            {


                new SelectListItem{ Text="ROS", Value = "ros",Selected=false },
                new SelectListItem{ Text="AOD", Value = "aod",Selected=false },
                new SelectListItem{ Text="Other", Value = "other",Selected=false },
                                         };
            return returnstatus;

        }

        private List<SelectListItem> GetPageNEFeilds()
        {

            List<SelectListItem> returnstatus = new List<SelectListItem>
            {


                new SelectListItem{ Text="Neurological Exam", Value = "neurological_exam",Selected=false },
                new SelectListItem{ Text="Sensory", Value = "sensory",Selected=false },
                new SelectListItem{ Text="Manual muscle strength testing", Value = "manual_muscle_strength_testing",Selected=false },
                new SelectListItem{ Text="DTR", Value = "other_content",Selected=false },
                                         };
            return returnstatus;

        }
        private List<SelectListItem> GetPageOtherFeilds()
        {

            List<SelectListItem> returnstatus = new List<SelectListItem>
            {
                new SelectListItem{ Text="Treatment", Value = "treatment_details",Selected=false },
                new SelectListItem{ Text="FOLLOW UP", Value = "followup_duration",Selected=false },
                new SelectListItem{ Text="FOLLOW Date", Value = "followup_date",Selected=false },

            };
            return returnstatus;

        }

        private List<SelectListItem> GetDaignosisFeilds()
        {

            List<SelectListItem> returnstatus = new List<SelectListItem>
            {
                new SelectListItem{ Text="Gait", Value = "gait",Selected=false },
                new SelectListItem{ Text="Goal", Value = "goal",Selected=false },
                new SelectListItem{ Text="Care", Value = "care",Selected=false },
                new SelectListItem{ Text="Universal", Value = "universal",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge Date", Value = "diagcervialbulge_date",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge Study", Value = "diagcervialbulge_study",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge", Value = "diagcervialbulge",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge Comma", Value = "diagcervialbulge_comma",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge Text", Value = "diagcervialbulge_text",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge hnp1", Value = "diagcervialbulge_hnp1",Selected=false },
                new SelectListItem{ Text="Diagcervialbulge hnp2", Value = "diagcervialbulge_hnp2",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge Date", Value = "diagthoracicbulge_date",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge Study", Value = "diagthoracicbulge_study",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge", Value = "diagthoracicbulge",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge Comma", Value = "diagthoracicbulge_comma",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge Text", Value = "diagthoracicbulge_text",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge_hnp1", Value = "diagthoracicbulge_hnp1",Selected=false },
                new SelectListItem{ Text="Diagthoracicbulge hnp2", Value = "diagthoracicbulge_hnp2",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge Date", Value = "diaglumberbulge_date",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge Study", Value = "diaglumberbulge_study",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge", Value = "diaglumberbulge",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge Comma", Value = "diaglumberbulge_comma",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge Text", Value = "diaglumberbulge_text",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge hnp1", Value = "diaglumberbulge_hnp1",Selected=false },
                new SelectListItem{ Text="Diaglumberbulge hhp2", Value = "diaglumberbulge_hnp2",Selected=false },
                new SelectListItem{ Text="Diagleftshoulder Date", Value = "diagleftshoulder_date",Selected=false },
                new SelectListItem{ Text="Diagleftshoulder Study", Value = "diagleftshoulder_study",Selected=false },
                new SelectListItem{ Text="Diagleftshoulder", Value = "diagleftshoulder",Selected=false },
                new SelectListItem{ Text="Diagleftshoulder Comma", Value = "diagleftshoulder_comma",Selected=false },
                new SelectListItem{ Text="Diagleftshoulder Text", Value = "diagleftshoulder_text",Selected=false },
                new SelectListItem{ Text="Diagrightshoulder Date", Value = "diagrightshoulder_date",Selected=false },
                new SelectListItem{ Text="Diagrightshoulder Study", Value = "diagrightshoulder_study",Selected=false },
                new SelectListItem{ Text="Diagrightshoulder", Value = "diagrightshoulder",Selected=false },
                new SelectListItem{ Text="Diagrightshoulder Comma", Value = "diagrightshoulder_comma",Selected=false },
                new SelectListItem{ Text="Diagrightshoulder Text", Value = "diagrightshoulder_text",Selected=false },
                new SelectListItem{ Text="Diagleftknee Date", Value = "diagleftknee_date",Selected=false },
                new SelectListItem{ Text="Diagleftknee Study", Value = "diagleftknee_study",Selected=false },
                new SelectListItem{ Text="Diagleftknee", Value = "diagleftknee",Selected=false },
                new SelectListItem{ Text="Diagleftknee Comma", Value = "diagleftknee_comma",Selected=false },
                new SelectListItem{ Text="Diagleftknee Text", Value = "diagleftknee_text",Selected=false },
                new SelectListItem{ Text="Diagrightknee Date", Value = "diagrightknee_date",Selected=false },
                new SelectListItem{ Text="Diagrightknee Study", Value = "diagrightknee_study",Selected=false },
                new SelectListItem{ Text="Diagrightknee Text", Value = "diagrightknee_text",Selected=false },
                new SelectListItem{ Text="Diagbrain Date", Value = "diagbrain_date",Selected=false },
                new SelectListItem{ Text="Diagbrain Study", Value = "diagbrain_study",Selected=false },
                new SelectListItem{ Text="Diagbrain", Value = "diagbrain",Selected=false },
                new SelectListItem{ Text="Diagbrain Comma", Value = "diagbrain_comma",Selected=false },
                new SelectListItem{ Text="Diagbrain Text", Value = "diagbrain_text",Selected=false },
                new SelectListItem{ Text="Other1 Date", Value = "other1_date",Selected=false },
                new SelectListItem{ Text="Other1 Study", Value = "other1_study",Selected=false },
                new SelectListItem{ Text="Other1", Value = "other1",Selected=false },
                new SelectListItem{ Text="Other1 Comma", Value = "other1_comma",Selected=false },
                new SelectListItem{ Text="Other1 Text", Value = "other1_text",Selected=false },
                new SelectListItem{ Text="Other2 Date", Value = "other2_date",Selected=false },
                new SelectListItem{ Text="Other2 Study", Value = "other2_study",Selected=false },
                new SelectListItem{ Text="Other2", Value = "other2",Selected=false },
                new SelectListItem{ Text="Other2 Comma", Value = "other2_comma",Selected=false },
                new SelectListItem{ Text="Other2 Text", Value = "other2_text",Selected=false },
                new SelectListItem{ Text="Other3 Date", Value = "other3_date",Selected=false },
                new SelectListItem{ Text="Other3 Study", Value = "other3_study",Selected=false },
                new SelectListItem{ Text="Other3", Value = "other3",Selected=false },
                new SelectListItem{ Text="Other3 Comma", Value = "other3_comma",Selected=false },
                new SelectListItem{ Text="Other3 Text", Value = "other3_text",Selected=false },
                new SelectListItem{ Text="Other4 Date", Value = "other4_date",Selected=false },
                new SelectListItem{ Text="Other4 Study", Value = "other4_study",Selected=false },
                new SelectListItem{ Text="Other4 Comma", Value = "other4_commma",Selected=false },
                new SelectListItem{ Text="Other4 Text", Value = "other4_text",Selected=false },
                new SelectListItem{ Text="Other5 Date", Value = "other5_date",Selected=false },
                new SelectListItem{ Text="Other5 Study", Value = "other5_study",Selected=false },
                new SelectListItem{ Text="Other5", Value = "other5",Selected=false },
                new SelectListItem{ Text="Other5 Comma", Value = "other5_comma",Selected=false },
                new SelectListItem{ Text="Other5 Text", Value = "other5_text",Selected=false },
                new SelectListItem{ Text="Other6 Date", Value = "other6_date",Selected=false },
                new SelectListItem{ Text="Other6 Study", Value = "other6_study",Selected=false },
                new SelectListItem{ Text="Other6 Comma", Value = "other6_comma",Selected=false },
                new SelectListItem{ Text="Other6 Text", Value = "other6_text",Selected=false },
                new SelectListItem{ Text="Other7 Date", Value = "other7_date",Selected=false },
                new SelectListItem{ Text="Other7 Study", Value = "other7_study",Selected=false },
                new SelectListItem{ Text="Other7", Value = "other7",Selected=false },
                new SelectListItem{ Text="Other7 Comma", Value = "other7_comma",Selected=false },
                new SelectListItem{ Text="Other7 Text", Value = "other7_text",Selected=false },
                new SelectListItem{ Text="Dollowupin", Value = "followupin",Selected=false },
                new SelectListItem{ Text="Followupin Date", Value = "followupin_date",Selected=false },
                new SelectListItem{ Text="Discharge Medications", Value = "discharge_medications",Selected=false },

            };
            return returnstatus;

        }

        public List<SelectListItem> GetVisitType()
        {
            string cnd = " and type<>'IE'";
            var data = _visitTypeService.GetAll(cnd);
            var list = new List<SelectListItem>();



            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.type,
                    Value = item.type.ToString()
                });
            }

            return list;
        }

        public List<SelectListItem> GetVisitTypeForTemplate(string cmpid)
        {
            string cnd = " and type<>'IE' and type not in (select type from tbl_template where cmp_id=" + cmpid + ")";



            var data = _visitTypeService.GetAll(cnd);
            var list = new List<SelectListItem>();



            foreach (var item in data)
            {
                list.Add(new SelectListItem
                {
                    Text = item.type,
                    Value = item.type.ToString()
                });
            }

            return list;
        }

        public string GetBodyPart(string bodyPart = "")
        {
            if (bodyPart.ToLower() == "cervical")
                return "Neck";
            if (bodyPart.ToLower() == "thoracic")
                return "Midback";
            if (bodyPart.ToLower() == "lumbar")
                return "Lowback";

            return bodyPart;
        }
    }


    public class CommonSp
    {
        #region Common
        //Common
        public const string CheckExecuteStatus = "sp_check_execute_status";
        public const string GetAllProceduress = "sp_GetAllProceduress";
        public const string SavePatientProceduresDetails = "sp_SavePatientProceduresDetails";
        public const string SavePatientProceduresDetailsBHF = "sp_SavePatientProceduresDetailsBHF";
        #endregion
    }


}
