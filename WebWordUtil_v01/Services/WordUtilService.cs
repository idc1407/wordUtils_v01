using System;
using System.IO;
using System.Threading.Tasks;
using System.Web.Hosting;
using WebWordUtil_v01.Models;

namespace WebWordUtil_v01.Services
{
    public static class WordUtilService
    {
        public static async Task<string> ProcessWordDocument(WordUtilWebModel wordUtilWebModel)
        {
            string status = "";
            try
            {
                var file = wordUtilWebModel.File;
                if (file.ContentLength > 0)
                {
                    string path = HostingEnvironment.MapPath("~/App_Data");

                    string sourcePath = Path.Combine(path, Path.GetFileName(file.FileName));
                    string targetPath = Path.Combine(path, "temp2.docx");

                    file.SaveAs(sourcePath);

                    WordUtilLib.Models.WordUtilModel wordUtilModel = new WordUtilLib.Models.WordUtilModel
                    {
                        SourceFileName = sourcePath,
                        TargetFileName = targetPath,
                        IsFooterTextChange = wordUtilWebModel.IsFooterTextChange,
                        FooterTextFind = wordUtilWebModel.FooterTextFind,
                        FooterTextReplace = wordUtilWebModel.FooterTextReplace,
                        IsHeaderTextChange = wordUtilWebModel.IsHeaderTextChange,
                        HeaderTextFind = wordUtilWebModel.HeaderTextFind,
                        HeaderTextReplace = wordUtilWebModel.HeaderTextReplace,
                        IsBalanceSheetTableDelete = wordUtilWebModel.IsBalanceSheetTableDelete,
                        IsOtherOptionA = wordUtilWebModel.IsOtherOptionA,
                        IsOtherOptionB = wordUtilWebModel.IsOtherOptionB
                    };
                    status = await WordUtilLib.Main.Process(wordUtilModel);
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