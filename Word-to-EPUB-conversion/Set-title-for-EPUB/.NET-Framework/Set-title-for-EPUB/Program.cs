using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_title_for_EPUB
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads the existing Word document by using DocIO instance.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Sets title for Word document, which will be applied as title for the output EPUB file.
                document.BuiltinDocumentProperties.Title = "This is a title in EPub document";
                //Saves and closes the document.
                document.Save(Path.GetFullPath(@"../../Result.epub"), FormatType.EPub);
            }
        }
    }
}
