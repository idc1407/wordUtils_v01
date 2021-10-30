using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using WebWordUtil_v01.Models;
using WebWordUtil_v01.Services;

namespace WebWordUtil_v01.Controllers
{
    public class UploadController : Controller
    {

        [HttpGet]
        public ActionResult UploadFile()
        {
            UploadFileModel uploadModel = new UploadFileModel
            {
                FooterTextFind = "Footer text" ,
                FooterTextReplace = "All is good"
            };
            
            return View(uploadModel);
        }


        

        [HttpPost]
        public async Task<ActionResult> UploadFile(UploadFileModel model)
        {
            if (!ModelState.IsValid)
            {
                ViewBag.Message = "Invalid Data.";
                return View();
            }

            string status = await WordUtilService.ProcessWordDocument(model);
            if(String.IsNullOrEmpty(status))
            {
                ViewBag.Message = "Job Completed Successfully!!";
            }
            else
            {
                ViewBag.Message = "*** Job Failed. ***";
            }
            return View();
        }
    }
}