using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Convert_Word_to_EPUB
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Exports the fonts used in the document.
                wordDocument.SaveOptions.EPubExportFont = true;
                //Exports header and footer.
                wordDocument.SaveOptions.HtmlExportHeadersFooters = true;
                //Saves the document as EPub file.
                wordDocument.Save(Path.GetFullPath(@"../../WordToEPub.epub"), FormatType.EPub);
            }
        }
    }
}
