using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

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
            test11();
        }


        public static string test11()
        {
            try
            {
                Application wordApp = new Application { Visible = true };
                Document doc = wordApp.Documents.Open(@"d:\itemp\test_para.docx", ReadOnly: false, Visible: true);


                foreach (Section section in doc.Sections)
                {
                    Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    headerRange.Text = "Header text goes here";
                }

                foreach (Section wordSection in doc.Sections)
                {
                    Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = "Footer text goes here";
                }



            }
            catch (Exception EX)
            {

                Console.WriteLine(EX.ToString());
            }
            return "";
        }





        public static string test10()
        {
            try
            {
                Application wordApp = new Application { Visible = true };
                Document doc = wordApp.Documents.Add();


                Range range = doc.Content;
                range.Text = "Hello world!";

                range.SetRange(Start: doc.Range().End, End: doc.Range().End);

                range.Text = "Bye for now!";

                // doc.Content.Select();
                range.Select();
                //range = doc.Content;


            }
            catch (Exception EX)
            {

                Console.WriteLine(EX.ToString());
            }
            return "";
        }




        public static string test09()
        {
            object fileName_01 = @"d:\itemp\test_para.docx";
            Application wordApp = new Application { Visible = false };

            Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: true);
            object Unknown = Type.Missing;

            Range rangeFind;
            

            List<Range> lst = new List<Range>() ;

            string status = "";
            try
            {
                rangeFind = wordDoc01.Range(0, 0);
                Find find = rangeFind.Find;
                find.ClearFormatting();
                find.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                while (find.Execute())
                {
                    lst.Add(rangeFind.Duplicate);
                }

                foreach (var item in lst){
                    item.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
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
            return status;
        }


        public static string test08()
        {
            object fileName_01 = @"d:\itemp\test_para.docx";
            Application wordApp = new Application { Visible = false };

            Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: true);
            object Unknown = Type.Missing;

            Range rangeFind;
            Range rangeFound;

            string status = "";
            try
            {
                rangeFind = wordDoc01.Range(0, 0);
                Find find = rangeFind.Find;
                find.ClearFormatting();
                find.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                while (find.Execute())
                {
                    rangeFound = rangeFind.Duplicate;
                    rangeFound.InsertBefore("<H1>");
                    rangeFound.MoveEnd(Unit: WdUnits.wdCharacter, Count: -1);
                    rangeFound.InsertAfter("<B1>");
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
            return status;
        }


        public static string test07()
        {
            try
            {
                Application wordApp = new Application { Visible = false };
                object missing = System.Reflection.Missing.Value;
                Document doc = wordApp.Documents.Add();


                Range range = doc.Content;
                range.Text = "Hello world!";

                range.InsertParagraphAfter();
                range = doc.Paragraphs.Last.Range;

                // start of list
                int startOfList = range.Start;

                // each \n character adds a new paragraph...
                range.Text = "Item 1\nItem 2\nItem 3";

                // ...or insert a new paragraph...
                range.InsertParagraphAfter();
                range = doc.Paragraphs.Last.Range;
                range.Text = "Item 4\nItem 5";

                // end of list
                int endOfList = range.End;

                // insert the next paragraph before applying the format, otherwise
                // the format will be copied to the suceeding paragraphs.
                range.InsertParagraphAfter();

                // apply list format
                Range listRange = doc.Range(startOfList, endOfList);
                listRange.ListFormat.ApplyBulletDefault();

                range = doc.Paragraphs.Last.Range;
                range.Text = "Bye for now!";
                range.InsertParagraphAfter();



                string path = Environment.CurrentDirectory;
                int totalExistDocx = Directory.GetFiles(path, "test*.docx").Count();
                path = Path.Combine(path, string.Format("test{0}.docx", totalExistDocx + 1));

                wordApp.ActiveDocument.SaveAs2(path);
                doc.Close();

                Process.Start(path);
            }
            catch (Exception )
            {

                throw;
            }
            return "";
        }

        public static string test06()
        {
            object fileName = @"d:\itemp\temp2.docx";
            Application wordApp = new Application { Visible = false };
            Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            

            string status = "";
            try
            {

                Range hrange = wordDoc.Range(0,0);
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
            Application wordApp = new Application { Visible = false };
            Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                Range hrange = wordDoc.Range(0, 0);
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
            Application wordApp = new Application { Visible = false };

            Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: false);
            Document wordDoc02 = wordApp.Documents.Open(fileName_02, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {

                wordDoc01.Activate();
                //wordApp.Selection.GoTo(What: WdGoToItem.wdGoToPage, Count: 1);
                wordApp.Selection.EndKey(Unit: WdUnits.wdStory, Extend: WdMovementType.wdExtend);
                //wordApp.Selection.Copy();

                wordDoc01.Activate();
                wordDoc01.Content.Select();
                wordApp.Selection.Copy();


                wordDoc02.Activate();
                wordApp.Selection.HomeKey(Unit: WdUnits.wdStory);
                wordApp.Selection.Find.Text = "test";

                while (wordApp.Selection.Find.Execute())
                {
                    if(wordApp.Selection.Font.Size == 20)
                    {
                        //wordApp.Selection.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

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
            Application wordApp = new Application { Visible = false };

            Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc01.Activate();

                foreach (Paragraph objParagraph in wordDoc01.Paragraphs)
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
            Application wordApp = new Application { Visible = false };

            Document wordDoc01 = wordApp.Documents.Open(fileName_01, ReadOnly: false, Visible: false);
            Document wordDoc02 = wordApp.Documents.Open(fileName_02, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {



                wordDoc01.Activate();
                wordDoc01.Content.Select();
                wordApp.Selection.MoveEnd(WdUnits.wdParagraph, 1);
                wordApp.Selection.Copy();

                wordDoc02.Activate();
                wordApp.Selection.Paste();


                //var docRange = wordDoc02.Content;
                //wordDoc02.Application.Selection.Find.ClearFormatting();
                //Find findObject = docRange.Find;
                //findObject.Text = "hhiugh";
                //findObject.Forward = true;
                //findObject.Execute();
                //if (findObject.Found)
                //{
                //    docRange.Expand(WdUnits.wdParagraph);
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

            Application wordApp = new Application { Visible = false };
            object missing = System.Reflection.Missing.Value;

            Document wordDoc = wordApp.Documents.Add();

            try
            {
                wordDoc.Content.Text = "ivan";

                wordDoc.Content.Select();
                wordApp.Selection.Copy();

                wordApp.Selection.EndKey(WdUnits.wdStory, missing);

                
                //wordApp.Selection.GoTo(
        //What: WdGoToItem.wdGoToPage,
        //Which: WdGoToDirection.wdGoToAbsolute,
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
                Find findObject = docRange.Find;
                findObject.Text = "rum";
                findObject.Forward = true;
                //findObject.Execute();
                //if (findObject.Found)
                //{
                //    docRange.Expand(WdUnits.wdParagraph);
                //    docRange.Delete();
                //}


                while (true)
                {
                    findObject.Execute();
                    if (!findObject.Found)
                    {
                        break;
                    }
                    docRange.Expand(WdUnits.wdParagraph);
                    float x = docRange.Font.Size;
                    docRange.Delete();
                }

                object filename = @"d:\itemp\t1.docx";


                wordDoc.SaveAs2(ref filename);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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
            Application wordApp = new Application { Visible = false };
            Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();
                object replaceAll = WdReplace.wdReplaceAll;
                foreach (Section section in wordDoc.Sections)
                {
                    HeadersFooters footers = section.Footers;
                    foreach (HeaderFooter footer in footers)
                    {
                        Range footerRange = footer.Range;
                        footerRange.Find.ClearFormatting();
                        footerRange.Find.Replacement.ClearFormatting();
                        footerRange.Find.Text = findText;
                        footerRange.Find.Replacement.Text = replaceText;
                        footerRange.Find.Wrap = WdFindWrap.wdFindContinue;
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
            Application wordApp = new Application { Visible = false };
            Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();
                Range hrange = wordDoc.Range(0, 0);

                hrange.Find.Execute("Range finder");

                object missing = System.Reflection.Missing.Value;
                object ConfirmConversions = false;
                object Link = false;
                object Attachment = false;

                hrange.Select();

                hrange.Bookmarks["\\Page"].Range.Delete();
                hrange.InsertFile(@"d:\itemp\test.docx", ref missing, ref ConfirmConversions, ref Link, ref Attachment);

                //hrange.Find.Execute("This is test document ivan"); 

                //hrange.InsertBreak(WdBreakType.wdPageBreak);


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
            Application wordApp = new Application { Visible = false };
            Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();


                wordDoc.Select();



                wordDoc.Application.Selection.Find.ClearFormatting();
                Find findObject = wordDoc.Application.Selection.Find;
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
            Application wordApp = new Application { Visible = false };
            Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            object Unknown = Type.Missing;

            string status = "";
            try
            {
                wordDoc.Activate();
                Range range = wordDoc.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "Para 1 text", ReplaceWith: "RRPara 1 text", Replace: WdReplace.wdReplaceAll);

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

