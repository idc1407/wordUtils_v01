using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordUtilTest
{
    class Program
    {
        static void Main(string[] args)
        {

            using (WordprocessingDocument doc =
          WordprocessingDocument.Open(@"d:\itemp\Test.docx", false))
            {
                foreach (var cc in doc.ContentControls())
                {
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    Tag tag = props.Elements<Tag>().FirstOrDefault();
                    Console.WriteLine(tag.Val);
                }
            }

            Console.ReadKey();
        }


    }



    public static class ContentControlExtensions
    {
        public static IEnumerable<OpenXmlElement> ContentControls(
            this OpenXmlPart part)
        {
            return part.RootElement
                .Descendants()
                .Where(e => e is SdtBlock || e is SdtRun);
        }

        public static IEnumerable<OpenXmlElement> ContentControls(
            this WordprocessingDocument doc)
        {
            foreach (var cc in doc.MainDocumentPart.ContentControls())
                yield return cc;
            foreach (var header in doc.MainDocumentPart.HeaderParts)
                foreach (var cc in header.ContentControls())
                    yield return cc;
            foreach (var footer in doc.MainDocumentPart.FooterParts)
                foreach (var cc in footer.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.FootnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.EndnotesPart != null)
                foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
                    yield return cc;
        }
    }



}

