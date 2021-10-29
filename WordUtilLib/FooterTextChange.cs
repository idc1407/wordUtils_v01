﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinWord = Microsoft.Office.Interop.Word;

namespace WordUtilLib
{
    public static class FooterTextChange
    {
        public static string Process(WinWord.Document wordDoc, string findText, string replaceText)
        {
            string status = "";
            object Unknown = Type.Missing;
            try
            {

                object replaceAll = WinWord.WdReplace.wdReplaceAll;
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
            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            return status;
        }


    }
}