using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Add_checkbox_using_IF_field
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the template document
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Merge field names and values
                    string[] fieldName = { "Name", "Email", "Option1","Option2" };
                    string[] fieldValue = { "John", "john@gmail.com", "YES","NO" };
                    //Execute mail merge
                    document.MailMerge.Execute(fieldName, fieldValue);
                    //Update fields
                    document.UpdateDocumentFields();

                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }

                }
            }
        }
    }
}
