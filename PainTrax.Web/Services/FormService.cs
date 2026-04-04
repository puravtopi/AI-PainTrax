namespace PainTrax.Web.Services
{
    using Newtonsoft.Json;
    using PainTrax.Web.Models;

    public class FormService
    {
        public List<FormField> GetFieldsFromJson(string jsonPath)
        {
            string jsonContent = System.IO.File.ReadAllText(jsonPath);

            return JsonConvert.DeserializeObject<List<FormField>>(jsonContent);
        }

    }
}
