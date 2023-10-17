using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using static System.Collections.Specialized.BitVector32;


namespace Customize_Word_to_HTML_conversion
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document into DocIO instance.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Rtf))
                {

                    //The headers and footers in the input are exported
                    document.SaveOptions.HtmlExportHeadersFooters = true;
                    //Exports the text form fields as editable 
                    document.SaveOptions.HtmlExportTextInputFormFieldAsText = false;
                    //Sets the style sheet type
                    document.SaveOptions.HtmlExportCssStyleSheetType = CssStyleSheetType.Inline;
                    //Set value to omit XML declaration in the exported html file.
                    //True- to omit xml declaration, otherwise false.
                    document.SaveOptions.HtmlExportOmitXmlDeclaration = false;

                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../RtfToHTML.html"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Html);
                    }
                }
            }
        }
    }
}
