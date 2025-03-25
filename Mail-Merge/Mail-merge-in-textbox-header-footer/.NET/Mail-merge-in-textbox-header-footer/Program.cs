using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Mail_merge_in_textbox_header_footer
{
    class Program
    {
        static void Main(string[] args)
        {

            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    string[] fieldNames = new string[] { "HeaderContent", "ProductName1", "ProductName2" };
                    string[] fieldValues = new string[] { "Adventure Works Cycles", "Mountain-200", "Mountain-300" };
                    //Performs the mail merge
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
