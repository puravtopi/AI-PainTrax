using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace PainTrax.Web.Helper
{
    public class SessionCheckFilter : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            var session = context.HttpContext.Session.GetString(SessionKeys.SessionUserId);
            if (string.IsNullOrEmpty(session))
            {
                context.Result = new RedirectToRouteResult(
                    new RouteValueDictionary
                    {
                    { "controller", "Home" },
                    { "action", "Login" }
                    });
            }
            base.OnActionExecuting(context);
        }
    }
}
