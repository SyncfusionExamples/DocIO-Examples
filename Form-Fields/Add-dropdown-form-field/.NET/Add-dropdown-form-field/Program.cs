using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_dropdown_form_field
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
                paragraph.AppendText("Educational Qualification\t");
                //Appends Dropdown field.
                WDropDownFormField dropDownField = paragraph.AppendDropDownFormField();
                //Adds items to the Dropdown items collection.
                dropDownField.DropDownItems.Add("Higher");
                dropDownField.DropDownItems.Add("Vocational");
                dropDownField.DropDownItems.Add("Universal");
                dropDownField.Enabled = true;
                //Sets the item index for default value.
                dropDownField.DropDownSelectedIndex = 1;
                dropDownField.CalculateOnExit = true;
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
