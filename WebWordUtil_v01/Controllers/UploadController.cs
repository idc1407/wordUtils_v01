using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebWordUtil_v01.Models;
using WordUtilLib;

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
        public ActionResult UploadFile(UploadFileModel model)
        {
            if (!ModelState.IsValid)
            {
                return View();
            }
            
            try
            {
                var file = model.File;
                bool x = model.IsFooterTextChange;

                if (file.ContentLength > 0)
                {
                    string _FileName = Path.GetFileName(file.FileName);
                    string _path = Path.Combine(Server.MapPath("~/UploadedFiles"), _FileName);
                    file.SaveAs(_path);

                    string destFileName = Path.Combine(Server.MapPath("~/UploadedFiles"), "temp2.docx");


                    string[] textReplce = { model.FooterTextFind, model.FooterTextReplace };

                    WordUtilLib.Main.Process(
                        _path,
                        destFileName,
                        textReplce
                        );
                }
                ViewBag.Message = "Job Completed Successfully!!";
                return View();
            }
            catch (Exception)
            {
                ViewBag.Message = "*** Job Failed. ***";
                return View();
            }
        }
    }
}