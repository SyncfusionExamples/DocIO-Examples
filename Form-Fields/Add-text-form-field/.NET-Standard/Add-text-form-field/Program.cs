using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Add_text_form_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                WParagraph paragraph = section.AddParagraph() as WParagraph;
                paragraph.AppendText("General Information");
                section.AddParagraph();
                paragraph = section.AddParagraph() as WParagraph;
                IWTextRange text = paragraph.AppendText("Name\t");
                text.CharacterFormat.Bold = true;
                //Appends Text form field.
                WTextFormField textField = paragraph.AppendTextFormField(null);
                //Sets type of Text form field.
                textField.Type = TextFormFieldType.RegularText;
                textField.CharacterFormat.FontName = "Calibri";
                textField.CalculateOnExit = true;
                section.AddParagraph();
                paragraph = section.AddParagraph() as WParagraph;
                text = paragraph.AppendText("Date of Birth\t");
                text.CharacterFormat.Bold = true;
                //Appends Text form field.
                textField = paragraph.AppendTextFormField("Date field", DateTime.Now.ToShortDateString());
                textField.StringFormat = "MM/DD/YY";
                //Sets Text form field type.
                textField.Type = TextFormFieldType.DateText;
                textField.CalculateOnExit = true;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
