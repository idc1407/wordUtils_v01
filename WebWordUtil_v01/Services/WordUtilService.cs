using System;
using System.IO;
using System.Threading.Tasks;
using System.Web.Hosting;
using WebWordUtil_v01.Models;

namespace WebWordUtil_v01.Services
{
    public static class WordUtilService
    {
        public static async Task<string> ProcessWordDocument(WordUtilModel uploadFileModel)
        {
            string status = "";
            try
            {
                var file = uploadFileModel.File;
                if (file.ContentLength > 0)
                {
                    string path = HostingEnvironment.MapPath("~/App_Data");
                    
                    string sourcePath = Path.Combine(path, Path.GetFileName(file.FileName));
                    string targetPath = Path.Combine(path, "temp2.docx");

                    file.SaveAs(sourcePath);

                    string[] textReplce = { uploadFileModel.FooterTextFind, uploadFileModel.FooterTextReplace };

                    status = await WordUtilLib.Main.Process(sourcePath, targetPath, textReplce);
                }
            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            return status;
        }
    }
}