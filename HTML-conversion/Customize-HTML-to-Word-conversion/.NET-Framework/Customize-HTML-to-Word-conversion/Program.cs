using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Customize_HTML_to_Word_conversion
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Html string to be inserted.
                string htmlstring = "<p><b>This text is inserted as HTML string.</b></p>";
                //Validates the Html string.
                bool isValidHtml = wordDocument.LastSection.Body.IsValidXHTML(htmlstring, XHTMLValidationType.Transitional);
                //When the Html string passes validation, it is inserted to the document.
                if (isValidHtml)
                {
                    //Appends Html string as first item of the second paragraph in the document.
                    wordDocument.Sections[0].Body.InsertXHTML(htmlstring, 2, 0);
                    //Appends the Html string to first paragraph in the document.
                    wordDocument.Sections[0].Body.Paragraphs[0].AppendHTML(htmlstring);
                }
                //Saves and closes the document.
                wordDocument.Save(Path.GetFullPath(@"../../Result.docx"));
            }
        }
    }
}
