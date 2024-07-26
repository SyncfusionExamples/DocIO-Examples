using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Drop_down_list_and_combo_box
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
                //Adds text to the paragraph.
                paragraph.AppendText("Choose your platform: ");
                //Appends dropdown list content control to the paragraph.
                InlineContentControl dropdown = paragraph.AppendInlineContentControl(ContentControlType.DropDownList) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                //Sets default option to display.
                textRange.Text = "Choose an item";
                dropdown.ParagraphItems.Add(textRange);
                //Creates an item for dropdown list.
                ContentControlListItem item = new ContentControlListItem();
                //Sets the text to be displayed as list item.
                item.DisplayText = "ASP.NET MVC";
                //Sets the value to the list item.
                item.Value = "1";
                //Adds item to the dropdown list.
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Windows Forms";
                item.Value = "2";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "WPF";
                item.Value = "3";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                //Adds new paragraph to the section.
                paragraph = section.AddParagraph() as WParagraph;
                //Adds text to the paragraph.
                paragraph.AppendText("Choose the conversion: ");
                //Appends combo box content control to the paragraph.
                InlineContentControl comboBox = paragraph.AppendInlineContentControl(ContentControlType.ComboBox) as InlineContentControl;
                textRange = new WTextRange(document);
                //Sets default option to display. 
                textRange.Text = "Choose an item";
                comboBox.ParagraphItems.Add(textRange);
                //Creates an item for combo box.
                item = new ContentControlListItem();
                //Sets the text to be displayed as list item.
                item.DisplayText = "Word to HTML";
                //Sets the value to the list item.
                item.Value = "1";
                comboBox.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Word to Image";
                item.Value = "2";
                comboBox.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Word to PDF";
                item.Value = "3";
                comboBox.ContentControlProperties.ContentControlListItems.Add(item);
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
