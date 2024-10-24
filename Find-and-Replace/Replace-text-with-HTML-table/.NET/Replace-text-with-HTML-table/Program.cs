using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_text_with_HTML_table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
                {
                    using (FileStream htmlFileStream = new FileStream(Path.GetFullPath(@"Data/Table.html"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open an HTML document.
                        using (WordDocument htmlDocument = new WordDocument(htmlFileStream, FormatType.Html))
                        {
                            //Get the first table from the HTML document.
                            WTable table = htmlDocument.Sections[0].Tables[0].Clone() as WTable;
                            TextBodyPart bodyPart = new TextBodyPart(document);
                            //Add the table to the body part.
                            bodyPart.BodyItems.Add(table);
                            //Replace the placeholder text "<<Product Table>>" with the HTML table in the Word document.
                            document.Replace("<<Product Table>>", bodyPart, true, true, false);
                            //Create a file stream to save the modified Word document.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the modified Word document to the file stream.
                                document.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}