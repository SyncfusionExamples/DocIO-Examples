
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
                //Adds a header paragraph to the section.
                IWParagraph headerPar = section.AddParagraph();
                //Appends some text to the header paragraph in document.
                headerPar.AppendText("Header text");
                // Add the header paragraph to the header section.
                section.HeadersFooters.Header.Paragraphs.Add(headerPar);
                //Adds a footer paragraph to the section.
                IWParagraph footerPar = section.AddParagraph();
                //Appends some text to the footer paragraph in document.
                footerPar.AppendText("Footer text");
                //Add the footer paragraph to the Footer section. 
                section.HeadersFooters.Footer.Paragraphs.Add(footerPar);
                //Save the PDF file.
                wordDocument.Save(Path.GetFullPath(@"../../Result.docx"));
            }
        }
    }
}
