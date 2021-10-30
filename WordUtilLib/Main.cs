using System;
using System.Threading.Tasks;
using WinWord = Microsoft.Office.Interop.Word;
using WordUtilLib.Models;

namespace WordUtilLib
{
    public static class Main
    {
        public static async Task<string> Process(WordUtilModel wordUtilModel)
        {
            string status = "";

            object fileName = wordUtilModel.SourceFileName;
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;
            wordDoc.Activate();
            try
            {
                if (wordUtilModel.IsFooterTextChange)
                {
                    status = await Utils.FooterTextReplace(wordDoc, wordUtilModel.FooterTextFind, wordUtilModel.FooterTextReplace);
                }
                
                if (status == "" && wordUtilModel.IsHeaderTextChange)
                {
                    status = await Utils.HeaderTextReplace(wordDoc, wordUtilModel.HeaderTextFind, wordUtilModel.HeaderTextReplace);
                }

                if (status == "" && wordUtilModel.IsBalanceSheetTableDelete)
                {
                    //TODO
                    status = "";
                }

                if (status == "" && wordUtilModel.IsOtherOptionA)
                {
                    //TODO
                    status = "";
                }

                if (status == "" && wordUtilModel.IsOtherOptionB)
                {
                    //TODO
                    status = "";
                }


                if (status == "")
                {
                    object filename = wordUtilModel.TargetFileName;
                    wordDoc.SaveAs(ref filename);
                }
            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            finally
            {
                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit(ref Unknown, ref Unknown, ref Unknown);
                wordApp = null;
            }
            return status;
        }
    }
}
