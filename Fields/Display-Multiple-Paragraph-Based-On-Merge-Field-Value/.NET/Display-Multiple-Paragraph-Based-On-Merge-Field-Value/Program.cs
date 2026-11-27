using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Display_Multiple_Paragraph_Based_On_Merge_Field_Value
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Word template document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
            {
                // Define merge field names available in the template.
                string[] fieldNames = new string[] { "VariableField1", "VariableField2" };

                // Define values to be merged into the corresponding fields.
                string[] fieldValues = new string[] { "1", "2" };

                // Perform mail merge using the specified field names and values.
                document.MailMerge.Execute(fieldNames, fieldValues);

                // Update all document fields after mail merge.
                document.UpdateDocumentFields();

                // Save the generated document as a DOCX file.
                document.Save(Path.GetFullPath(@"Output/Output.docx"), FormatType.Docx);
            }
        }
    }
}
