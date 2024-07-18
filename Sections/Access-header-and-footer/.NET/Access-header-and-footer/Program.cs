using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Access_header_and_footer
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens the Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Adds header and footer to each section in the document.
                    foreach (WSection sec in document.Sections)
                    {
                        //Header.
                        WParagraph headerParagraph = new WParagraph(document);
                        headerParagraph.AppendField("page", FieldType.FieldPage);
                        headerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                        sec.HeadersFooters.Header.Paragraphs.Add(headerParagraph);
                        //Footer.
                        WParagraph footerParagraph = new WParagraph(document);
                        footerParagraph.AppendText("Internal");
                        footerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                        sec.HeadersFooters.Footer.Paragraphs.Add(footerParagraph);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
