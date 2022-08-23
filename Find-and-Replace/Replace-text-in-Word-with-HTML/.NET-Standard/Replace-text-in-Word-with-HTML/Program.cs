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
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Create the temporary word document for html.
                    using (WordDocument replaceDoc = new WordDocument())
                    {
                        //Add section for html document. 
                        IWSection htmlsection = replaceDoc.AddSection();
                        //Reads HTML string from the file.
                        string htmlString = File.ReadAllText(Path.GetFullPath(@"../../../Data/File.html"));
                        //Validates the Html string.
                        bool isValidHtml = htmlsection.Body.IsValidXHTML(htmlString, XHTMLValidationType.None);
                        //When the Html string passes validation, it is inserted to the document.
                        if (isValidHtml)
                        {
                            //Appends Html string in the temporary word document.
                            htmlsection.Body.InsertXHTML(htmlString);
                        }
                        //Replaces the content placeholder text with desired Word document.
                        document.Replace(new Regex("«([(?i)image(?-i)]*:*[a-zA-Z0-9 ]*:*[a-zA-Z0-9 ]+)»"), replaceDoc, true);
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
