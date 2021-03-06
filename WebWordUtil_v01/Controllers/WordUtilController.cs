using System;
using System.Threading.Tasks;
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
            WordUtilWebModel wordUtilWebModel = new WordUtilWebModel
            {
                FooterTextFind = "Footer text" ,
                FooterTextReplace = "All is good"
            };
            return View(wordUtilWebModel);
        }
       

        [HttpPost]
        public async Task<ActionResult> ProcessFile(WordUtilWebModel wordUtilWebModel)
        {
            ViewBag.ErrorMessage = "";
            if (!ModelState.IsValid)
            {
                ViewBag.Message = "Invalid Data.";
                return View();
            }

            string status = await WordUtilService.ProcessWordDocument(wordUtilWebModel);
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