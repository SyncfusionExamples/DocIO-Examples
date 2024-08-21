using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Document_variables
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of a WordDocument.
            using (WordDocument document = new WordDocument())
            {
                //Adds a new section into the Word Document.
                IWSection section = document.AddSection();
                //Adds a new paragraph into Word document and appends text into paragraph.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendText("First Name of the customer: ");
                //Adds the DocVariable field with Variable name and its type.
                paragraph.AppendField("FirstName", FieldType.FieldDocVariable);
                paragraph = section.AddParagraph();
                paragraph.AppendText("Last Name of the customer: ");
                //Adds the DocVariable field with Variable name and its type.
                paragraph.AppendField("LastName", FieldType.FieldDocVariable);
                //Adds the value for variable in WordDocument.Variable collection.
                document.Variables.Add("FirstName", "Jeff");
                document.Variables.Add("LastName", "Smith");
                //Updates the document fields.
                document.UpdateDocumentFields();
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
