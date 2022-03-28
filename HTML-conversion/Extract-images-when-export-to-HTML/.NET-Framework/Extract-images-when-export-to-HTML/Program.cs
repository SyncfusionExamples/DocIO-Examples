using Syncfusion.DocIO.DLS;
using System.IO;

namespace Extract_images_when_export_to_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads the template document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Template.docx")))
            {
                //Sets the location to extract images.
                document.SaveOptions.HtmlExportImagesFolder = @"D:\Images\";
                //Saves the document as html file.
                HTMLExport export = new HTMLExport();
                export.SaveAsXhtml(document, Path.GetFullPath(@"../../Result.html"));
            }
        }
    }
}
