using Microsoft.AspNetCore.Mvc;
using MS.Services;
using PainTrax.Web.Filter;
using PainTrax.Web.Helper;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class AutoCompleteController : Controller
    {

        int? cmp_id = 0;


        [HttpPost]
        public JsonResult InsCoAutoComplete(string prefix)
        {
            InscosService _inscosService = new InscosService();
            cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var cnd = " and cmp_id=" + cmp_id.Value + " and cmpname like '%" + prefix + "%'";
            var customers = _inscosService.GetAautoComplete(cnd);

            return Json(customers);
        }

         [HttpPost]
        public JsonResult StateAutoComplete(string prefix)
        {
            StateService _stateService = new StateService();
            cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var cnd = " and state_name like '%" + prefix + "%'";
            var state = _stateService.GetAautoComplete(cnd);

            return Json(state);
        }

        [HttpPost]
        public JsonResult AttornyAutoComplete(string prefix)
        {
            AttorneysService _attyService = new AttorneysService();
            cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var cnd = " and cmp_id=" + cmp_id.Value + " and attorney like '%" + prefix + "%'";
            var attorny = _attyService.GetAautoComplete(cnd);

            return Json(attorny);
        }


        [HttpPost]
        public JsonResult AdjusterAutoComplete(string prefix)
        {
            AadjusterService _adjService = new AadjusterService();
            cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var cnd = " and cmp_id=" + cmp_id.Value + " and adjuster like '%" + prefix + "%'";
            var attorny = _adjService.GetAautoComplete(cnd);

            return Json(attorny);
        }


        [HttpPost]
        public JsonResult GlobalAutoComplete(string prefix)
        {
           WebsiteMacrosMasterService _adjService = new WebsiteMacrosMasterService();
            cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
            var cnd = " and t.cmp_id=" + cmp_id.Value + " and t.type='All' and t.key like '%" + prefix + "%'";
            var attorny = _adjService.GetAautoComplete(cnd);

            return Json(attorny);
        }
    }
}
