using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_checkbox_form_field
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
                paragraph.AppendText("Gender : ");
                //Appends new Checkbox.
                WCheckBox checkbox = paragraph.AppendCheckBox();
                checkbox.Checked = false;
                //Sets Checkbox size.
                checkbox.CheckBoxSize = 10;
                checkbox.CalculateOnExit = true;
                //Sets help text.
                checkbox.Help = "Help text";
                paragraph.AppendText("Male\t");
                checkbox = paragraph.AppendCheckBox();
                checkbox.Checked = false;
                checkbox.CheckBoxSize = 10;
                checkbox.CalculateOnExit = true;
                paragraph.AppendText("Female");
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
