using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Date_picker_content_control
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds one section and one paragraph to the document.
                document.EnsureMinimal();
                //Gets the last paragraph.
                WParagraph paragraph = document.LastParagraph;
                //Adds text to the paragraph.
                paragraph.AppendText("Select Date: ");
                //Appends date picker content control to the paragraph.
                InlineContentControl datePicker = paragraph.AppendInlineContentControl(ContentControlType.Date) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                //Sets today's date to display.
                textRange.Text = DateTime.Now.ToShortDateString();
                datePicker.ParagraphItems.Add(textRange);
                //Sets calendar type for the date picker content control.
                datePicker.ContentControlProperties.DateCalendarType = CalendarType.Gregorian;
                //Sets the format for date to display.
                datePicker.ContentControlProperties.DateDisplayFormat = "M/d/yyyy";
                //Sets the language format for the date.
                datePicker.ContentControlProperties.DateDisplayLocale = LocaleIDs.en_US;
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
