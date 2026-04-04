using Azure;
using Azure.AI.Vision.ImageAnalysis;
using Microsoft.Extensions.Options;
using System.Text.RegularExpressions;
namespace PainTrax.Web.AzureServices
{
    public class AzureAIServices
    {
        private readonly AzureVisionSettings _settings;

        public AzureAIServices(IOptions<AzureVisionSettings> settings)
        {
            _settings = settings.Value;
        }
        public async Task<string> ExtractTextFromImage(string imagePath)
        {
            var endpoint = _settings.Endpoint;
            var key = _settings.Key;

            var client = new ImageAnalysisClient(
                new Uri(endpoint),
                new AzureKeyCredential(key));

            using var stream = File.OpenRead(imagePath);

            var result = await client.AnalyzeAsync(
                BinaryData.FromStream(stream),
                VisualFeatures.Read);

            var text = "";

            foreach (var block in result.Value.Read.Blocks)
            {
                foreach (var line in block.Lines)
                {
                    text += line.Text + "\n";
                }
            }

            return text;
        }

        public LicenseData ParseData(string text)
        {
            var result = new LicenseData();

            // ================= DOB =================
            var dobMatch = Regex.Match(text, @"DOB[:\s]*([0-9\-\/]+)");
            //var dobMatch = Regex.Match(text, @"DOB\s*[:\-]?\s*([0-9]{2}[\/\-][0-9]{2}[\/\-][0-9]{4})", RegexOptions.IgnoreCase);

            if (dobMatch.Success)
                result.DOB = dobMatch.Groups[1].Value;

            // ================= Gender =================
            var genderMatch = Regex.Match(text, @"S\s*[E]?\s*X\s*[:\-]?\s*([MF])", RegexOptions.IgnoreCase);

            if (genderMatch.Success)
                result.Gender = genderMatch.Groups[1].Value == "M" ? "male" : "female";

            // ================= Address =================
            var addressMatch = Regex.Match(text, @"\d{3,5}.*\n.*\d{5}");
            if (addressMatch.Success)
                result.Address = addressMatch.Value.Replace("\n", " ");

            // ================= Name Extraction =================

            var lines = text.Split('\n')
                            .Select(l => l.Trim())
                            .Where(l => !string.IsNullOrEmpty(l))
                            .ToList();

            for (int i = 0; i < lines.Count; i++)
            {
                // Case 1: Normal or broken lastname with comma
                if (lines[i].Contains(","))
                {
                    var parts = lines[i].Split(',');

                    if (parts.Length >= 2)
                    {
                        string rawLastName = parts[0].Trim();
                        string rawFirstName = parts[1].Trim();

                        // 🔥 Handle case: "SMITH," + next line "JOHN"
                        if (string.IsNullOrWhiteSpace(rawFirstName) && i + 1 < lines.Count)
                        {
                            rawFirstName = lines[i + 1].Trim();
                        }

                        // 🔥 Handle multi-line lastname (previous line)
                        if (i > 0 && !lines[i - 1].Contains(","))
                        {
                            rawLastName = lines[i - 1].Trim() + " " + rawLastName;
                        }

                        // 🔥 Clean OCR noise
                        result.FirstName = CleanName(rawFirstName);
                        result.LastName = CleanName(rawLastName);

                        break;
                    }
                }
            }
            return result;
        }

        private string CleanName(string input)
        {
            return Regex.Replace(input, @"[^A-Za-z\s]", "").Trim();
        }
    }
}
