using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Customize_Word_to_HTML_conversion
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                HTMLExport export = new HTMLExport();
                //The images in the input document are copied to this folder.
                wordDocument.SaveOptions.HtmlExportImagesFolder = @"D:/Images/";
                //The headers and footers in the input are exported.
                wordDocument.SaveOptions.HtmlExportHeadersFooters = true;
                //Exports the text form fields as editables.
                wordDocument.SaveOptions.HtmlExportTextInputFormFieldAsText = false;
                //Sets the style sheet type.
                wordDocument.SaveOptions.HtmlExportCssStyleSheetType = CssStyleSheetType.External;
                //Sets name for style sheet.
                wordDocument.SaveOptions.HtmlExportCssStyleSheetFileName = "UserDefinedFileName.css";
                //Export the Word document image as Base-64 embedded image.
                wordDocument.SaveOptions.HTMLExportImageAsBase64 = false;
                //Saves the document as html file.
                export.SaveAsXhtml(wordDocument, Path.GetFullPath(@"../../WordtoHtml.html"));
            }
        }
    }
}
