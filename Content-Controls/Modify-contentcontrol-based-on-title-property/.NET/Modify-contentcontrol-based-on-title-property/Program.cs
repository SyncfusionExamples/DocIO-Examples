using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Modify_contentcontrol_based_on_tag_property
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
                    //Finds all the inline content controls in the Word document
                    List<Entity> inlineContentControls = document.FindAllItemsByProperty(EntityType.InlineContentControl, null, null);
                    //Iterates through the inline content controls
                    foreach (Entity entity in inlineContentControls)
                    {
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        //Replaces the content when the content control title is "ReplaceText"
                        if (inlineContentControl.ContentControlProperties.Title == "ReplaceText")
                            ReplaceTextWithInlineContentControl("Hello World", inlineContentControl);
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
