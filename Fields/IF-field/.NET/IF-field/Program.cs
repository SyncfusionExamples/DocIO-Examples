using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace IF_field
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
                paragraph.AppendText("If field which uses string of characters in expression");
                paragraph = section.AddParagraph();
                //Creates the new instance of IF field.
                WIfField field = paragraph.AppendField("If", FieldType.FieldIf) as WIfField;
                //Specifies the expression, true and false statement in field code.
                field.FieldCode = "IF \"True\" = \"True\" \"The given statement is Correct\" \"The given statement is Wrong\"";
                paragraph = section.AddParagraph();
                paragraph.AppendText("If field which uses numbers in expression");
                paragraph = section.AddParagraph();
                //Creates the new instance of IF field.
                field = paragraph.AppendField("If", FieldType.FieldIf) as WIfField;
                //Specifies the expression, true and false statement in field code.
                field.FieldCode = "IF 100 >= 1000 \"The given statement is Correct\" \"The given statement is Wrong\"";
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
