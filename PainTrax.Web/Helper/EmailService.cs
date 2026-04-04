using MailKit.Net.Smtp;
using MimeKit;

namespace PainTrax.Web.Helper
{
    public class EmailService : IEmailService
    {
        private readonly IConfiguration _configuration;

        public EmailService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public async Task SendEmailAsync(string toEmail, string subject, string body)
        {
            var emailSettings = _configuration.GetSection("SmtpSettings");

            var email = new MimeMessage();
            email.From.Add(new MailboxAddress(emailSettings["SenderEmail"], emailSettings["SenderEmail"]));
            email.To.Add(MailboxAddress.Parse(toEmail));
            email.Subject = subject;
            email.Body = new TextPart("html") { Text = body };

            using var smtp = new SmtpClient();
            await smtp.ConnectAsync(emailSettings["Server"], int.Parse(emailSettings["Port"]), MailKit.Security.SecureSocketOptions.StartTls);
            await smtp.AuthenticateAsync(emailSettings["SenderEmail"], emailSettings["SenderPassword"]);
            await smtp.SendAsync(email);
            await smtp.DisconnectAsync(true);
        }
    }
}