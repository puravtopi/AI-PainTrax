using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using PainTrax.Web.Helper;
using PainTrax.Web.Models;
using PainTrax.Web.Services;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;
using System;
using Optivem.Framework.Core.Domain;
using System.Diagnostics;
using Org.BouncyCastle.Asn1.Cmp;

namespace PainTrax.Web.Controllers
{
    [SessionCheckFilter]
    public class SignatureController : Controller
    {
        private readonly ILogger<SignatureController> _logger;
        private readonly IMapper _mapper;
        private readonly SignatureService _services;
        private readonly IConfiguration _configuration;
        private readonly Microsoft.AspNetCore.Hosting.IHostingEnvironment Environment;

        public SignatureController(ILogger<SignatureController> logger, IMapper mapper, SignatureService service, IConfiguration configuration, Microsoft.AspNetCore.Hosting.IHostingEnvironment environment)
        {
            _logger = logger;
            _mapper = mapper;
            _services = service;
            _configuration = configuration;
            Environment = environment;
        }

        public IActionResult Index()
        {
           // var model = _services.GetAll(); 
            return View();
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create([FromBody] tbl_signature model)
        {
            
            model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);

            try
            {           

                var base64Data = model.signaturePath.Contains(",")
                                ? model.signaturePath.Split(',')[1]
                                : model.signaturePath;

                if (!IsBase64String(base64Data))
                {
                    return BadRequest("Invalid Base-64 string.");
                }

                var imageData = Convert.FromBase64String(base64Data);

                var signaturesDir = Path.Combine(Environment.WebRootPath, "signatures", model.cmp_id.ToString());
                if (!Directory.Exists(signaturesDir))
                {
                    Directory.CreateDirectory(signaturesDir);
                }               

                var dobFormatted = DateTime.ParseExact(model.dob, "yyyy-MM-dd", null).ToString("yyyy-MM-dd");
                model.dob = dobFormatted;
                var filename = $"{model.lname}_{model.fname}_{model.dob}.jpeg";
                var savePath = Path.Combine(signaturesDir, filename);

                System.IO.File.WriteAllBytes(savePath, imageData);
                model.signaturePath = savePath;
                _services.Insert(model);
                return View("Index");

                // return Ok(new { FileName = filename });

            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error saving signature: " + ex.Message);
            }
        }


        public IActionResult Edit(int id)
        {
            var data = _services.GetOne(id);
            return View(data);
        }

        [HttpPost]
        public IActionResult Edit(tbl_signature model)
        {
            try
            {
                model.cmp_id = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId);
                var base64Data = model.signaturePath.Contains(",")
                                ? model.signaturePath.Split(',')[1]
                                : model.signaturePath;

                if (!IsBase64String(base64Data))
                {
                    return BadRequest("Invalid Base-64 string.");
                }

                var imageData = Convert.FromBase64String(base64Data);

                var signaturesDir = Path.Combine(Environment.WebRootPath, "signatures", model.cmp_id.ToString());
                if (!Directory.Exists(signaturesDir))
                {
                    Directory.CreateDirectory(signaturesDir);
                }

                var dobFormatted = DateTime.ParseExact(model.dob, "yyyy-MM-dd", null).ToString("yyyy-MM-dd");
                model.dob = dobFormatted;
                var filename = $"{model.lname}_{model.fname}_{model.dob}.jpeg";
                var savePath = Path.Combine(signaturesDir, filename);

                System.IO.File.WriteAllBytes(savePath, imageData);
                model.signaturePath = savePath;
                _services.Update(model);
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Edit");
            }
            return View(model);
        }

        public IActionResult Delete(int id)
        {
            try
            {
                _services.Delete(new tbl_signature { id = id });
                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                SaveLog(ex, "Delete");
            }
            return RedirectToAction("Index");
        }
      
        [HttpPost]
        public IActionResult List()
        {
            try
            {
                var cmpid = HttpContext.Session.GetInt32(SessionKeys.SessionCmpId).ToString();
                var draw = HttpContext.Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["order[0][column]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var searchValue = Request.Form["search[value]"].FirstOrDefault();

                int pageSize = length != null ? Convert.ToInt32(length) : 0;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                var cnd = " and cmp_id=" + cmpid;

                var data = _services.GetAll(cnd);

                if (!string.IsNullOrEmpty(sortColumn) && !string.IsNullOrEmpty(sortColumnDirection))
                {
                    var property = typeof(tbl_signature).GetProperties()[Convert.ToInt32(sortColumn)];
                    data = sortColumnDirection.ToUpper() == "ASC"
                        ? data.OrderBy(x => property.GetValue(x, null)).ToList()
                        : data.OrderByDescending(x => property.GetValue(x, null)).ToList();
                }

                recordsTotal = data.Count();
                var result = data.Skip(skip).Take(pageSize).ToList();

                return Json(new { draw, recordsFiltered = recordsTotal, recordsTotal, data = result });
            }
            catch (Exception ex)
            {
                SaveLog(ex, "List");
                return Json(new { error = ex.Message });
            }
        }

        


        private bool IsBase64String(string base64)
        {
            Span<byte> buffer = new Span<byte>(new byte[base64.Length]);
           return Convert.TryFromBase64String(base64, buffer, out _);
        }

        #region Private Methods
        private void SaveLog(Exception ex, string actionName)
        {
            var msg = ex.InnerException?.Message ?? ex.Message;
            _logger.LogError(msg);
            var logData = new tbl_log
            {
                CreatedDate = DateTime.Now,
                CreatedBy = HttpContext.Session.GetInt32(SessionKeys.SessionCmpUserId),
                Message = msg
            };
            new LogService().Insert(logData);
        }
        #endregion
    }
}
