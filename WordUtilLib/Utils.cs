using System;
using System.Threading.Tasks;
using WinWord = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace WordUtilLib
{
    public static class Utils
    {
        public static async Task<string> XmlFooterTextReplace(string sourcePath, string findText, string replaceText)
        {
            string status = "";
            try
            {
                await Task.Run(() =>
                {
                    using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(sourcePath, true))
                    {
                        foreach (var footerPart in wdDoc.MainDocumentPart.FooterParts)
                        {
                            foreach (var currentText in footerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                            {
                                currentText.Text = currentText.Text.Replace(findText, replaceText);
                            }
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


        public static async Task<string> FooterTextReplace(WinWord.Document wordDoc, string findText, string replaceText)
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


        public static async Task<string> HeaderTextReplace(WinWord.Document wordDoc, string findText, string replaceText)
        {
            string status = "";
            object Unknown = Type.Missing;
            try
            {
                object replaceAll = WinWord.WdReplace.wdReplaceAll;
                await Task.Run(() =>
                {
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
