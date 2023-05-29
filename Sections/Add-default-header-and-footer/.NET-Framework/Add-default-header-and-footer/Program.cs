
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_default_header_and_footer
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new document.
            using (WordDocument wordDocument = new WordDocument())
            {
                // Add a new section to the document.
                IWSection section = wordDocument.AddSection();
                //Adds a paragraph to the header.
                IWParagraph headerPar = section.HeadersFooters.Header.AddParagraph();
                //Appends some text to the paragraph in document.
                headerPar.AppendText("Header text");
                //Adds a paragraph to the footer.
                IWParagraph footerPar = section.HeadersFooters.Footer.AddParagraph();
                //Appends some text to the paragraph in document.
                footerPar.AppendText("Footer text");
                //Save the Word document.
                wordDocument.Save(Path.GetFullPath(@"../../Result.docx"));
            }
        }
    }
}
