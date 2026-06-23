using Microsoft.AspNetCore.Mvc;
using MS.Services;
using Newtonsoft.Json.Linq;
using Org.BouncyCastle.Asn1.Ocsp;
using PainTrax.Services;
using System.Data;
using System.Text.RegularExpressions;

namespace PainTrax.Web.Controllers
{

    public class JsonController : Controller
    {
        private readonly ILogger<FormsController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly PatientIEService _ieService = new PatientIEService();
        private readonly PatientService _patientservices = new PatientService();
        private readonly ParentService _pareentservices = new ParentService();

        public JsonController(ILogger<FormsController> logger, IWebHostEnvironment environment)
        {
            _environment = environment;
            _logger = logger;
        }
        public static string ReplacePlaceholders(string template, List<Dictionary<string, string>> data)
        {
            if (data == null || data.Count == 0)
                return template;

            // Assuming there is only one record in the list
            Dictionary<string, string> values = data[0];

            return Regex.Replace(template, @"#(.*?)#", match =>
            {
                string key = match.Groups[1].Value;

                if (values.TryGetValue(key, out string value))
                {
                    return value ?? "";
                }

                // Replace missing key with empty string
                return "";
            });
        }
        public static List<Dictionary<string, string>> ExtractKeyValuePairs(string json)
        {
            var list = new List<Dictionary<string, string>>();
            var dict = new Dictionary<string, string>();

            JObject obj = JObject.Parse(json);

            foreach (var property in obj.Properties())
            {
                if (property.Value.Type == JTokenType.Array)
                {
                    dict[property.Name] = string.Join(", ", property.Value.ToObject<List<string>>() ?? new List<string>());
                }
                else
                {
                    dict[property.Name] = property.Value.ToString();
                }
            }

            list.Add(dict);
            return list;
        }

        public IActionResult Index()
        {
            int? intakeid = 74;
            string json = "";
            string message = "";
            DataTable data = null;
            var query = $"SELECT FormData FROM tbl_intake_ai WHERE id  ={intakeid} ";
            data = _pareentservices.GetData(query);
            if (data != null && data.Rows.Count > 0)
                   json = data.Rows[0]["FormData"].ToString();
            if (!string.IsNullOrEmpty(json))
            {
                var jsondata = ExtractKeyValuePairs(json);
                //foreach (var row in jsondata)
                //{
                //    message += "\n Record:";

                //    foreach (var item in row)
                //    {
                //        message += $"{item.Key} : {item.Value}";
                //    }
                //}

                string template = "The patient #FN# #LN# has complaints #Complaints#.";
                              

                message = ReplacePlaceholders(template, jsondata);

            }
            TempData["Message"] = message;
            return View();
        }
    }
}
