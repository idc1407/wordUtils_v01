using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using WinWord = Microsoft.Office.Interop.Word;

namespace WordUtilTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //FooterTextReplace("Footer text", "all is good");
            //insertPage();
            //FindText();
            //FindReplaceText();
            test06();
        }


        public static string test06()
        {
            object fileName = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                WinWord.Range hrange = wordDoc.Range(0,0);
                hrange.Select();
                Console.WriteLine(wordApp.Selection.Words.Count);
                Console.ReadKey();
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


        public static string test05()
        {
            object fileName = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                WinWord.Range hrange = wordDoc.Range(0, 0);
                while (hrange.Find.Execute("Range finder"))
                {
                 
                    if (hrange.Font.Size == 20)
                    {
                        object missing = System.Reflection.Missing.Value;
                        object ConfirmConversions = false;
                        object Link = false;
                        object Attachment = false;

                        hrange.Select();

                        hrange.Bookmarks["\\Page"].Range.Delete();
                        hrange.InsertFile(FileName: @"d:\itemp\test.docx", Range : missing, ConfirmConversions: ConfirmConversions, Link: Link, Attachment: Attachment);

                        break;
                    }

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




        public static string test04()
        {
            object fileName_01 = @"d:\itemp\test.docx";
            object fileName_02 = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };

            WinWord.Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: false);
            WinWord.Document wordDoc02 = wordApp.Documents.Open(fileName_02, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {

                wordDoc01.Activate();
                //wordApp.Selection.GoTo(What: WinWord.WdGoToItem.wdGoToPage, Count: 1);
                wordApp.Selection.EndKey(Unit: WinWord.WdUnits.wdStory, Extend: WinWord.WdMovementType.wdExtend);
                //wordApp.Selection.Copy();

                wordDoc01.Activate();
                wordDoc01.Content.Select();
                wordApp.Selection.Copy();


                wordDoc02.Activate();
                wordApp.Selection.HomeKey(Unit: WinWord.WdUnits.wdStory);
                wordApp.Selection.Find.Text = "test";

                while (wordApp.Selection.Find.Execute())
                {
                    if(wordApp.Selection.Font.Size == 20)
                    {
                        //wordApp.Selection.PageSetup.Orientation = WinWord.WdOrientation.wdOrientLandscape;

                        wordApp.Selection.Paste();
                    }
                }




            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            finally
            {
                wordDoc01.Close();
                wordDoc02.Close();
                wordDoc01 = null;
                wordDoc02 = null;
                wordApp.Quit(ref Unknown, ref Unknown, ref Unknown);
                wordApp = null;
            }
            return status;
        }

        public static void test03()
        {
            object fileName_01 = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };

            WinWord.Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc01.Activate();

                foreach (WinWord.Paragraph objParagraph in wordDoc01.Paragraphs)
                {
                    Console.WriteLine(objParagraph.Range.Text);
                }


            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            finally
            {
                wordDoc01.Close();
                wordDoc01 = null;
                wordApp.Quit(ref Unknown, ref Unknown, ref Unknown);
                wordApp = null;
            }
            Console.ReadKey();

        }


        public static string test02()
        {
            object fileName_01 = @"d:\itemp\test.docx";
            object fileName_02 = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };

            WinWord.Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: false);
            WinWord.Document wordDoc02 = wordApp.Documents.Open(fileName_02, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {



                wordDoc01.Activate();
                wordDoc01.Content.Select();
                wordApp.Selection.MoveEnd(WinWord.WdUnits.wdParagraph, 1);
                wordApp.Selection.Copy();

                wordDoc02.Activate();
                wordApp.Selection.Paste();


                //var docRange = wordDoc02.Content;
                //wordDoc02.Application.Selection.Find.ClearFormatting();
                //WinWord.Find findObject = docRange.Find;
                //findObject.Text = "hhiugh";
                //findObject.Forward = true;
                //findObject.Execute();
                //if (findObject.Found)
                //{
                //    docRange.Expand(WinWord.WdUnits.wdParagraph);
                //    docRange.Delete();
                //}



                var paragraphs = wordDoc02.Paragraphs;

                //if (paragraphs.Last.Range.Text.Trim() == string.Empty)
                //{
                paragraphs.Last.Range.Select();
                wordApp.Selection.Delete();
                //}

            }
            catch (Exception ex)
            {
                status = ex.ToString();
            }
            finally
            {
                wordDoc01.Close();
                wordDoc02.Close();
                wordDoc01 = null;
                wordDoc02 = null;
                wordApp.Quit(ref Unknown, ref Unknown, ref Unknown);
                wordApp = null;
            }
            return status;
        }


        public static void test01()
        {

            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            object missing = System.Reflection.Missing.Value;

            WinWord.Document wordDoc = wordApp.Documents.Add();

            try
            {
                wordDoc.Content.Text = "ivan";

                wordDoc.Content.Select();
                wordApp.Selection.Copy();

                wordApp.Selection.EndKey(WinWord.WdUnits.wdStory, missing);

                
                //wordApp.Selection.GoTo(
        //What: WinWord.WdGoToItem.wdGoToPage,
        //Which: WinWord.WdGoToDirection.wdGoToAbsolute,
        //Count: 3);
        


                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("tum tum tum");



                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("rum tum tum");


                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("rum tum tum");


                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("tum tum tum");

                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("rum tum tum");


                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("rum tum tum");



                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("rum tum tum");

                wordApp.Selection.InsertAfter(Environment.NewLine);
                wordApp.Selection.InsertAfter("tum tum tum");



                var docRange = wordDoc.Content;
                wordDoc.Application.Selection.Find.ClearFormatting();
                WinWord.Find findObject = docRange.Find;
                findObject.Text = "rum";
                findObject.Forward = true;
                //findObject.Execute();
                //if (findObject.Found)
                //{
                //    docRange.Expand(WinWord.WdUnits.wdParagraph);
                //    docRange.Delete();
                //}


                while (true)
                {
                    findObject.Execute();
                    if (!findObject.Found)
                    {
                        break;
                    }
                    docRange.Expand(WinWord.WdUnits.wdParagraph);
                    float x = docRange.Font.Size;
                    docRange.Delete();
                }

                object filename = @"d:\itemp\t1.docx";


                wordDoc.SaveAs2(ref filename);


            }
            catch (Exception ex)
            {

            }
            finally
            {
                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit();
                wordApp = null;

            }




        }

        public static string FooterTextReplace(string findText, string replaceText)
        {
            object fileName = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();
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
            finally
            {
                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit(ref Unknown, ref Unknown, ref Unknown);
                wordApp = null;
            }
            return status;
        }



        public static string insertPage()
        {
            object fileName = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();
                WinWord.Range hrange = wordDoc.Range(0, 0);

                hrange.Find.Execute("Range finder");

                object missing = System.Reflection.Missing.Value;
                object ConfirmConversions = false;
                object Link = false;
                object Attachment = false;

                hrange.Select();

                hrange.Bookmarks["\\Page"].Range.Delete();
                hrange.InsertFile(@"d:\itemp\test.docx", ref missing, ref ConfirmConversions, ref Link, ref Attachment);

                //hrange.Find.Execute("This is test document ivan"); 

                //hrange.InsertBreak(WinWord.WdBreakType.wdPageBreak);


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

        public static void FindText()
        {
            int cnt = 0;
            object fileName = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();


                wordDoc.Select();



                wordDoc.Application.Selection.Find.ClearFormatting();
                WinWord.Find findObject = wordDoc.Application.Selection.Find;
                findObject.Text = "Para 1 text";
                findObject.Forward = true;
                findObject.Execute();



                while (findObject.Found)
                {
                    cnt++;
                    findObject.Execute();
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
            Console.WriteLine(cnt.ToString());
            Console.ReadKey();

        }


        public static void FindReplaceText()
        {
            int cnt = 0;
            object fileName = @"d:\itemp\temp2.docx";
            WinWord.Application wordApp = new WinWord.Application { Visible = false };
            WinWord.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();
                WinWord.Range range = wordDoc.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "Para 1 text", ReplaceWith: "RRPara 1 text", Replace: WinWord.WdReplace.wdReplaceAll);

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
            Console.WriteLine(cnt.ToString());
            Console.ReadKey();

        }



    }
}

