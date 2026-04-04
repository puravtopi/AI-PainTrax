namespace PainTrax.Web.Helper
{
    public class ServiceResponse<T>
    {
        public T Model { get; set; }

        public bool Success { get; set; }

        public string Message { get; set; }

        public string ErrorCode { get; set; }

        public string jwtToken { get; set; }

        public string refreshToken { get; set; }

        public int? Id { get; set; }
    }
}
