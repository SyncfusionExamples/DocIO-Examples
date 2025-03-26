using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.IO;

namespace Mail_merge_Word_to_PDF
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Opens the template Word document
            using (WordDocument document = new WordDocument("../../Data/Template.docx"))
            {
                // Defines the merge field names and corresponding values
                string[] fieldNames = new string[] { "EmployeeId", "Name", "Phone", "City" };
                string[] fieldValues = new string[] { "1001", "Peter", "+122-2222222", "London" };

                // Performs the mail merge operation
                document.MailMerge.Execute(fieldNames, fieldValues);

                // Converts the merged Word document to PDF
                using (DocToPDFConverter render = new DocToPDFConverter())
                {
                    using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                    {
                        // Saves the PDF document to the specified file path
                        using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"../../Result.pdf"), FileMode.Create, FileAccess.Write))
                        {
                            pdfDocument.Save(docStream1);
                        }
                    }
                }
            }
        }
    }
}
