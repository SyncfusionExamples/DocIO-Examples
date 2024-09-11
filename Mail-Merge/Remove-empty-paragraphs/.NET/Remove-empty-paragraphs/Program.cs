using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_empty_paragraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Removes paragraph that contains only empty fields.
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    string[] fieldNames = new string[] { "EmployeeName", "EmployeeId", "City" };
                    string[] fieldValues = new string[] { "John", "101", "London" };
                    //Performs the mail merge.
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
