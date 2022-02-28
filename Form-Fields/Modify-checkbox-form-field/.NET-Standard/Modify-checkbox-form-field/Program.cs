using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_checkbox_form_field
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
                        if (item is WCheckBox)
                        {
                            WCheckBox checkbox = item as WCheckBox;
                            //Modifies check box properties.
                            if (checkbox.Checked)
                                checkbox.Checked = false;
                            checkbox.SizeType = CheckBoxSizeType.Exactly;
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
