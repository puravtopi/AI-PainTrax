using DocumentFormat.OpenXml.Packaging;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using PainTrax.Web.Models;
using System.Text;

namespace PainTrax.Web.Controllers
{
    public class AI_IntakeFormController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public async Task<IActionResult> Upload(FileUploadModel model)
        {
            if (model.File == null || model.File.Length == 0)
                return BadRequest("File not uploaded");

            string text = "";

            var extension = Path.GetExtension(model.File.FileName).ToLower();

            using (var stream = model.File.OpenReadStream())
            {
                if (extension == ".pdf")
                {
                    text = ExtractTextFromPdf(stream);
                }
                else if (extension == ".docx")
                {
                    text = ExtractTextFromDoc(stream);
                }
            }

            var questions = await GenerateQuestions(text);

            return Json(questions);

        }

        #region Private Method
        public string ExtractTextFromPdf(Stream stream)
        {
            StringBuilder text = new StringBuilder();

            using (PdfReader reader = new PdfReader(stream))
            using (PdfDocument pdf = new PdfDocument(reader))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(pdf.GetPage(i)));
                }
            }

            return text.ToString();
        }


        public string ExtractTextFromDoc(Stream stream)
        {
            StringBuilder text = new StringBuilder();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, false))
            {
                var body = doc.MainDocumentPart.Document.Body;

                foreach (var paragraph in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    text.AppendLine(paragraph.InnerText);
                }
            }

            return text.ToString();
        }

        public async Task<string> GenerateQuestions(string documentText)
        {
            documentText = documentText.Replace("L/R/B", " Left Right Both ");
            documentText = documentText.Replace("__ digits", "digits");
            var client = new HttpClient();

            string prompt = $@"
You are an intelligent document form analyzer.

Your job is to read the document and detect ALL possible form fields.

Detect fields including:
- Textboxes
- Textareas
- Checkboxes
- Radio buttons
- Dropdown selections
- Yes/No questions
- Multiple choice questions

Rules for detection:

1. If options like:
   ( ) Male   ( ) Female
   or
   ☐ Yes  ☐ No

   → type = """"radio"""" or """"checkbox""""

2. If multiple selectable items appear
   → type = """"checkbox""""

3. If one selection among many options
   → type = """"dropdown"""" or """"radio""""

4. If long answer space or paragraph
   → type = """"textarea""""

5. If simple input blank
   → type = """"text""""

6. Patterns like ___/10 represent a pain score from 0 to 10.
   Convert them to a numeric question.

7. Options like:
__Constant __Intermittent __Sharp __Electric __Shooting __Throbbing __Pulsating __Dull __Achy

represent multiple selectable pain types.
These should be converted into a checkbox question.

8. Each body part (Neck, Midback, Lowback, Shoulder, Knee) should become its own question group.

10. If a line lists body parts like:
Shoulder Scapula Arm Forearm Hand Wrist Digits

Treat them as checkbox options.

12. If the next line contains L/R/B it means side selection:
L = Left
R = Right
B = Both

13. Generate questions in structured JSON.


Return ONLY valid JSON.

JSON format:

[
 {{{{
   """"question"""": """"string"""",
   """"type"""": """"text | textarea | checkbox | radio | dropdown | number | date"""",
   """"options"""": [""""option1"""", """"option2""""]
 }}}}
]

Rules:
- If field has options include them in """"options"""" array
- If no options keep """"options"""": []

Example output:

[
 {{{{
   """"question"""": """"What is your name?"""",
   """"type"""": """"text"""",
   """"options"""": []
 }}}},
 {{{{
   """"question"""": """"Gender"""",
   """"type"""": """"radio"""",
   """"options"""": [""""Male"""", """"Female""""]
 }}}},
 {{{{
   """"question"""": """"Symptoms"""",
   """"type"""": """"textarea"""",
   """"options"""": []
 }}}},
 {{{{
   """"question"""": """"Do you smoke?"""",
   """"type"""": """"checkbox"""",
   """"options"""": [""""Yes"""", """"No""""]
 }}}}
{{
   ""question"": ""Rate your neck pain from 0 to 10"",
   ""type"": ""number"",
   ""options"": []
 }},
 {{
   ""question"": ""What type of neck pain do you experience?"",
   ""type"": ""checkbox"",
   ""options"": [""Constant"",""Intermittent"",""Sharp"",""Electric"",""Shooting"",""Throbbing"",""Pulsating"",""Dull"",""Achy""]
 }},
 {{
   ""question"": ""string"",
   ""type"": ""checkbox | radio | text | number"",
   ""options"": []
 }}
]

Document:
{documentText}
";

            var request = new
            {
                //model = "gpt-3.5-turbo",
                model = "gpt-4.1",
                messages = new[]
                {
            new {
                role = "user",
                content =prompt
            }
        }
            };

            

            var response = await client.PostAsJsonAsync(
                "https://api.openai.com/v1/chat/completions",
                request
            );

            var result = await response.Content.ReadAsStringAsync();


            var aiResponse = JsonConvert.DeserializeObject<AIResponse>(result);

            string jsonQuestions = aiResponse.choices[0].message.content;

            var qList = JsonConvert.DeserializeObject<List<FormField>>(jsonQuestions);

            return jsonQuestions;
        }

        #endregion
    }
}
