using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinWord = Microsoft.Office.Interop.Word;

namespace WordUtilLib
{
    public static class Main
    {

        public static string Process(string sourceFileName, string targetFileName, string[] options)
        {
            string status = "";


            object fileName = sourceFileName;
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;
            wordDoc.Activate();
            try
            {
                status = FooterTextChange.Process(wordDoc, options[0], options[1]);
                if (status == "")
                {
                    object filename = targetFileName;
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
