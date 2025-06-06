﻿using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Protect_content_control
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
                paragraph.AppendText("A new text is added to the paragraph. ");
                //Appends rich text content control to the paragraph.
                IInlineContentControl contentControl = paragraph.AppendInlineContentControl(ContentControlType.RichText) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                textRange.Text = "Rich text content control.";
                //Adds new text to the rich text content control.
                contentControl.ParagraphItems.Add(textRange);
                //Sets tag appearance for the content control.
                contentControl.ContentControlProperties.Appearance = ContentControlAppearance.Tags;
                //Sets a tag property to identify the content control.
                contentControl.ContentControlProperties.Tag = "Rich Text Protected";
                //Sets a title for the content control.
                contentControl.ContentControlProperties.Title = "Text Protected";
                //Enables content control lock.
                contentControl.ContentControlProperties.LockContentControl = true;
                //Protects the contents of content control.
                contentControl.ContentControlProperties.LockContents = true;
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
