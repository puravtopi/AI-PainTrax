namespace PainTrax.Web.Models
{
    public class FileUploadModel
    {
        public IFormFile File { get; set; }
    }

    public class AIResponse
    {
        public List<Choice> choices { get; set; }
    }

    public class Choice
    {
        public Message message { get; set; }
    }

    public class Message
    {
        public string content { get; set; }
    }

    public class FormField
    {
        public string Question { get; set; }
        public string Type { get; set; } // text, radio, checkbox, number, date, textarea
        public List<string> Options { get; set; } = new List<string>();
        public string Value { get; set; } // To store the user's input
        public List<string> SelectedOptions { get; set; } = new List<string>(); // For multi-select
    }
}
