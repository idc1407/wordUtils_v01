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
    public class WordUtilController : Controller
    {

        [HttpGet]
        public ActionResult ProcessFile()
        {
            WordUtilModel wordUtilModel = new WordUtilModel
            {
                FooterTextFind = "Footer text" ,
                FooterTextReplace = "All is good"
            };
            
            return View(wordUtilModel);
        }


        

        [HttpPost]
        public async Task<ActionResult> ProcessFile(WordUtilModel model)
        {
            ViewBag.ErrorMessage = "";
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

                ViewBag.ErrorMessage = status;
            }
            return View();
        }
    }
}