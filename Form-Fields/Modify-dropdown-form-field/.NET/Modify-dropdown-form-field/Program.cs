using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_dropdown_form_field
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Iterates through paragraph items.
                    foreach (ParagraphItem item in document.LastParagraph.ChildEntities)
                    {
                        if (item is WDropDownFormField)
                        {
                            WDropDownFormField dropdown = item as WDropDownFormField;
                            //Modifies the dropdown items.
                            dropdown.DropDownItems.Remove(1);
                            dropdown.DropDownSelectedIndex = 0;
                            dropdown.CharacterFormat.FontName = "Arial";
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
