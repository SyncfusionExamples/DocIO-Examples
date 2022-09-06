using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_in_Word_with_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Sample.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Create the temporary word document for HTML.
                    using (WordDocument replaceDoc = new WordDocument())
                    {
                        //Add section for HTML document. 
                        IWSection htmlsection = replaceDoc.AddSection();
                        //Read HTML string from the file.
                        string htmlString = File.ReadAllText(Path.GetFullPath(@"../../../Data/File.html"));
                        //Validate the HTML string.
                        bool isValidHtml = htmlsection.Body.IsValidXHTML(htmlString, XHTMLValidationType.None);
                        //When the HTML string passes validation, it is inserted to the document.
                        if (isValidHtml)
                        {
                            //Append HTML string in the temporary word document.
                            htmlsection.Body.InsertXHTML(htmlString);
                        }
                        //Replace the content placeholder text with desired Word document.
                        document.Replace(new Regex("«([a-zA-Z0-9 ]*:*[a-zA-Z0-9 ]+)»"), replaceDoc, true);
                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
