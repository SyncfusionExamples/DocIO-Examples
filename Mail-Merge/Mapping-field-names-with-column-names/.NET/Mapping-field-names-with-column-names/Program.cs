using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Mapping_field_names_with_column_names
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
                    //Creates data source.
                    string[] fieldNames = new string[] { "Employee_Id_InDataSource", "Name_InDataSource", "Phone_InDataSource", "City_InDataSource" };
                    string[] fieldValues = new string[] { "101", "John", "+122-2000466", "Houston" };
                    //Mapping the required merge field names with data source column names.
                    document.MailMerge.MappedFields.Add("Employee_Id_InDocument", "Employee_Id_InDataSource");
                    document.MailMerge.MappedFields.Add("Name_InDocument", "Name_InDataSource");
                    document.MailMerge.MappedFields.Add("Phone_InDocument", "Phone_InDataSource");
                    document.MailMerge.MappedFields.Add("City_InDocument", "City_InDataSource");
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
