using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Reflection.Metadata;

namespace Modify_contentcontrol_based_on_tag_property
{
    class program
    {
        static void Main(string[] args)
        {// Load the Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../../Data/Template.docx")))
            {
                // Retrieve all inline content controls from the document.
                List<Entity> inlineContentControls = document.FindAllItemsByProperty(EntityType.InlineContentControl, null, null);
                if (inlineContentControls != null)
                {
                    // Iterate through each content control and update its content.
                    foreach (Entity inlineContentControl in inlineContentControls)
                    {
                        GetContentControl(inlineContentControl as InlineContentControl);
                    }
                }
                document.Save(Path.GetFullPath(@"../../../Output/output.docx"));
            }
        }
        /// <summary>
        /// Identifies the inline content control based on its tag value
        /// and replaces its content with the corresponding text.
        /// </summary>
        /// <param name="inlineContentControl">
        /// The inline content control to be processed.
        /// </param>
        private static void GetContentControl(InlineContentControl inlineContentControl)
        {
            //Gets a tag property to identify the content control
            string tag = inlineContentControl.ContentControlProperties.Tag;
            switch (tag)
            {
                case "VideoTitle":
                    ReplaceTextWithInlineContentControl("AI Features Overview", inlineContentControl);
                    break;
                case "SearchKeyword":
                    ReplaceTextWithInlineContentControl("Productivity Tips", inlineContentControl);
                    break;
                case "DocumentName":
                    ReplaceTextWithInlineContentControl("Product Documentation", inlineContentControl);
                    break;
            }
        }
        /// <summary>
        /// Replaces the existing text within an inline content control
        /// while preserving the original character formatting.
        /// </summary>
        /// <param name="text">
        /// The replacement text to insert into the content control.
        /// </param>
        /// <param name="inlineContentControl">
        /// The target inline content control.
        /// </param>
        private static void ReplaceTextWithInlineContentControl(string text, InlineContentControl inlineContentControl)
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
            inlineContentControl.ParagraphItems.Clear();
            WTextRange textRange = new WTextRange(inlineContentControl.Document);
            textRange.Text = text;
            if (characterFormat != null)
                textRange.ApplyCharacterFormat(characterFormat);
            inlineContentControl.ParagraphItems.Add(textRange);
        }
    }
}

