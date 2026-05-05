using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Remove_InlineContentControls_Containing_PageField
{
    class Program
    {
        public static void Main(string[] args)
        {            
            //Load existing Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx"), FormatType.Docx))
            {
                //Find Content control
                List<Entity> entities = document.FindAllItemsByProperties(EntityType.InlineContentControl, null, null);
                // Iterate through the list of entities in reverse to safely remove items during iteration
                for (int i = entities.Count - 1; i >= 0; i--)
                {
                    // Attempt to cast the current entity to an InlineContentControl
                    InlineContentControl inLineContentControl = entities[i] as InlineContentControl;
                    if (inLineContentControl != null)
                    {                       
                        // Iterate through the child items of the content control
                        foreach (ParagraphItem entity in inLineContentControl.ParagraphItems)
                        {
                            // Check if the entity is a PAGE field
                            if (entity is WField field && field.FieldType == FieldType.FieldPage)
                            {
                                // If the content control is nested inside another content control
                                if (inLineContentControl.Owner is InlineContentControl)
                                {
                                    InlineContentControl owner = inLineContentControl.Owner as InlineContentControl;
                                    // Remove the current content control from its parent content control
                                    owner.ParagraphItems.Remove(inLineContentControl);
                                }
                                // If the content control is directly inside a paragraph
                                else if (inLineContentControl.Owner is WParagraph)
                                {
                                    WParagraph parentParagraph = inLineContentControl.Owner as WParagraph;
                                    // Remove the content control from the paragraph
                                    parentParagraph.ChildEntities.Remove(inLineContentControl);
                                }
                                // Remove the content control from the main entity list
                                entities.RemoveAt(i);
                                // Exit the inner loop since the control has been removed
                                break;
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