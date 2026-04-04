using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Optivem.Framework.Core.Domain;
using PainTrax.Web.Models;
using PainTrax.Web.Services;

namespace PainTrax.Web.Controllers
{
    public class TagsController : Controller
    {
        private readonly IMapper _mapper;
        private readonly ILogger<TagsController> _logger;
        private readonly TagServices _services = new TagServices();

        public TagsController(ILogger<TagsController> logger, IMapper mapper)
        {
            _logger = logger;
            _mapper = mapper;

        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult List()
        {
            try
            {

                var draw = HttpContext.Request.Form["draw"].FirstOrDefault();
                // Skiping number of Rows count
                var start = Request.Form["start"].FirstOrDefault();
                // Paging Length 10,20
                var length = Request.Form["length"].FirstOrDefault();
                // Sort Column Name
                var sortColumn = Request.Form["order[0][column]"].FirstOrDefault();
                // Sort Column Direction ( asc ,desc)
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                // Search Value from (Search box)
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                //Paging Size (10,20,50,100)
                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                string cnd = string.IsNullOrEmpty(searchValue) ? "" : " and  (tag like '%" + searchValue + "%' or `desc` like '%" + searchValue + "%')";
                var Data = _services.GetAll(cnd);

                //Sorting
                //if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                //{
                //    var property = typeof(tbl_state).GetProperties()[Convert.ToInt32(sortColumn)];
                //    if (sortColumnDirection.ToUpper() == "ASC")
                //    {
                //        Data = Data.OrderBy(x => property.GetValue(x, null)).ToList();
                //    }
                //    else
                //    {
                //        Data = Data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                //    }
                //}

                //Search

                //total number of rows count 
                recordsTotal = Data.Count();
                //Paging 
                var data = Data.Skip(skip).Take(pageSize).ToList();
                //Returning Json Data
                return Json(new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data });

            }
            catch (Exception ex)
            {
                // SaveLog(ex, "List");
            }
            return Json("");
        }
    }
}
