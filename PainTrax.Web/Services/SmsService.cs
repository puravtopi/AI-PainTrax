using Newtonsoft.Json;
using System.Text;

namespace PainTrax.Web.Services
{
    public class SmsService
    {
        private readonly HttpClient _httpClient;

        public SmsService(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<string> SendSmsAsync()
        {
            var url = "https://api.textedly.com/v1/send"; // confirm actual endpoint from dashboard

            var payload = new
            {
                target_numbers = new[] { "+9175824008" },
                content = "Hello from .NET Core!",
                wait_for_send = true
            };

            var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Add("Authorization", "Bearer 4b5b379b14b94343db52c5be5352bc9a143f4af8e90c3c1af47d406af9023f22");

            request.Content = new StringContent(
                JsonConvert.SerializeObject(payload),
                Encoding.UTF8,
                "application/json"
            );

            var response = await _httpClient.SendAsync(request);
            var result = await response.Content.ReadAsStringAsync();

            return result;
        }
    }
}
