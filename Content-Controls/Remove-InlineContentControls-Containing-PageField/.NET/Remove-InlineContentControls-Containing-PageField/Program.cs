using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Remove_InlineContentControls_Containing_PageField
{
    class Program
    {
        public static void Main(string[] args)
        {            
            // Load existing Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx"), FormatType.Docx))
            {
                // Find all PAGE fields in the document
                List<Entity> pageFields = document.FindAllItemsByProperty(EntityType.Field, "FieldType", "FieldPage");
                // Iterate through the list of entities in reverse to safely remove items during iteration
                for (int i = pageFields.Count - 1; i >= 0; i--)
                {
                    // Attempt to cast the current entity to an WField
                    WField pageField = pageFields[i] as WField;
                    if (pageField != null)
                    {
                        // Check the direct parent of PAGE field is InlineContentControl
                        if (pageField.Owner is InlineContentControl)
                        {
                            InlineContentControl inlineContentControl = pageField.Owner as InlineContentControl;
                            // If the content control is nested inside another content control
                            if (inlineContentControl.Owner is InlineContentControl)
                            {
                                InlineContentControl owner = inlineContentControl.Owner as InlineContentControl;
                                // Remove the current content control from its parent content control
                                owner.ParagraphItems.Remove(inlineContentControl);
                            }
                            // If the content control is directly inside a paragraph
                            else if (inlineContentControl.Owner is WParagraph)
                            {
                                WParagraph parentParagraph = inlineContentControl.Owner as WParagraph;
                                // Remove the content control from the paragraph
                                parentParagraph.ChildEntities.Remove(inlineContentControl);
                            }
                        }
                    }
                }               
                // Save the document if needed
                using (FileStream fileStream = new FileStream(@"../../../Output/Output.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    document.Save(fileStream, FormatType.Docx);
                }
            }
        }
    }
}