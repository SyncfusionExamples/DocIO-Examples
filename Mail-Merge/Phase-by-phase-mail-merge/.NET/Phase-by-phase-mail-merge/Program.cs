using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Phase_by_phase_mail_merge
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the input Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Disable the flag to maintain unmerged fields in document for next phase of execution.
                    document.MailMerge.ClearFields = false;

                    //First phase merge field execution.
                    string[] phase1_FieldName = new string[] { "EmployeeId" };
                    string[] phase1_FieldValue = new string[] { "1001" };
                    //Perform the mail merge.
                    document.MailMerge.Execute(phase1_FieldName, phase1_FieldValue);

                    //Second phase merge fields execution.
                    string[] phase2_FieldNames = new string[] { "Name", "Phone", "City" };
                    string[] phase2_FieldValues = new string[] { "Peter", "+122-2222222", "London" };
                    //Perform the mail merge.
                    document.MailMerge.Execute(phase2_FieldNames, phase2_FieldValues);

                    //Third phase merge field execution.
                    string[] phase3_FieldName = new string[] { "Country" };
                    string[] phase3_FieldValue = new string[] { "United Kingdom" };
                    //Perform the mail merge.
                    document.MailMerge.Execute(phase3_FieldName, phase3_FieldValue);
                    //Create file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
