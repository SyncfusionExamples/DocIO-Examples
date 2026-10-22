using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Edit_text_in_inline_content_control
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Creates an instance of WordDocument class
                using (WordDocument document = new WordDocument(docStream, FormatType.Automatic))
                {
                    ///Processes the body contents for each section in the Word document
                    // Use FindAllItemsByProperty to get all inline content controls with Title == "ReplaceText"
                    // and replace their text, instead of iterating through the DOM.
                    System.Collections.IList controls = document.FindAllItemsByProperty(EntityType.InlineContentControl, null, null);
                    if (controls != null)
                    {
                        foreach (object ctrl in controls)
                        {
                            InlineContentControl inlineContentControl = ctrl as InlineContentControl;
                            if (inlineContentControl != null && 
                                inlineContentControl.ContentControlProperties.Title == "ReplaceText")
                            {
                                ReplaceTextWithInlineContentControl("Hello World", inlineContentControl);
                            }
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Replace Text With Inline content control
        /// </summary>
        /// <param name="text"></param>
        /// <param name="inlineContentControl"></param>
        private static void ReplaceTextWithInlineContentControl(string replacementText, InlineContentControl inlineContentControl)
        {
            WCharacterFormat characterFormat = null;
            foreach (ParagraphItem item in inlineContentControl.ParagraphItems)
            {
                if (item is WTextRange)
                {
                    characterFormat = (item as WTextRange).CharacterFormat;
                    break;
                }
            }
            //Remove exiting items and add new text range with required text
            inlineContentControl.ParagraphItems.Clear();
            WTextRange textRange = new WTextRange(inlineContentControl.Document);
            textRange.Text = replacementText;
            if (characterFormat != null)
                textRange.ApplyCharacterFormat(characterFormat);
            inlineContentControl.ParagraphItems.Add(textRange);
        }
    }
}
