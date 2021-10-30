using System;
using System.Threading.Tasks;
using WinWord = Microsoft.Office.Interop.Word;

namespace WordUtilLib
{
    public static class FooterTextChange
    {
        public static async Task<string> Process(WinWord.Document wordDoc, string findText, string replaceText)
        {
            string status = "";
            object Unknown = Type.Missing;
            try
            {
                object replaceAll = WinWord.WdReplace.wdReplaceAll;
                await Task.Run(() =>
                {
                    foreach (WinWord.Section section in wordDoc.Sections)
                    {
                        WinWord.HeadersFooters footers = section.Footers;
                        foreach (WinWord.HeaderFooter footer in footers)
                        {
                            WinWord.Range footerRange = footer.Range;
                            footerRange.Find.ClearFormatting();
                            footerRange.Find.Replacement.ClearFormatting();
                            footerRange.Find.Text = findText;
                            footerRange.Find.Replacement.Text = replaceText;
                            footerRange.Find.Wrap = WinWord.WdFindWrap.wdFindContinue;
                            footerRange.Find.Execute(
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref replaceAll,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown,
                                ref Unknown
                                );
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            return status;
        }
    }
}
