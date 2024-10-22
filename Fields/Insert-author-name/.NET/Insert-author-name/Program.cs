using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_author_name
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Get the section in a Word document.
                    IWSection section = document.LastSection;
                    //Add paragraph to the document section.
                    IWParagraph paragraph = section.AddParagraph();
                    paragraph.AppendText("Author: ");
                    //Add field to represent Author from document properties.
                    paragraph.AppendField("Author", FieldType.FieldDocProperty);
                    //Update the fields in Word document.
                    document.UpdateDocumentFields();
                    //Create file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
