using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;

namespace PainTrax.Web.Filter
{
    public class AuthenticateUser : Attribute, IAuthorizationFilter
    {
        public void OnAuthorization(AuthorizationFilterContext context)
       {
            string tempSession =
                Convert.ToString(context.HttpContext.Session.GetString(SessionKeys.SessionCmpEmail));
          
            if (tempSession == null )
            {
                context.Result = (IActionResult)new RedirectToActionResult("SessionExpired", "Home",null);
            }
            
        }
    }
}
