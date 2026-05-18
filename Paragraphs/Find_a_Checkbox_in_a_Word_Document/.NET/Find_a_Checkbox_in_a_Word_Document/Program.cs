using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_a_Checkbox_in_a_Word_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            // Opens the input Word document from the specified path
            using (FileStream fileStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Loads the Word document into DocIO
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    // Finds all Block Content Controls of type CheckBox in the document
                    List<Entity> blockContentControls = document.FindAllItemsByProperty(EntityType.BlockContentControl, "ContentControlProperties.Type", ContentControlType.CheckBox.ToString());

                    // Verifies if any block-level checkbox content controls are found
                    if (blockContentControls != null)
                    {
                        foreach (Entity entity in blockContentControls)
                        {
                            // Cast the entity to BlockContentControl
                            BlockContentControl blockContentControl = entity as BlockContentControl;
                            // Unchecks the checkbox
                            blockContentControl.ContentControlProperties.IsChecked = false;
                        }
                    }

                    // Finds all Inline Content Controls of type CheckBox in the document
                    List<Entity> inlineContentControls = document.FindAllItemsByProperty(EntityType.InlineContentControl, "ContentControlProperties.Type", ContentControlType.CheckBox.ToString());

                    // Verifies if any inline checkbox content controls are found
                    if (inlineContentControls != null)
                    {
                        foreach (Entity entity in inlineContentControls)
                        {
                            // Cast the entity to InlineContentControl
                            InlineContentControl inlineContentControl = entity as InlineContentControl;
                            // Unchecks the checkbox
                            inlineContentControl.ContentControlProperties.IsChecked = false;
                        }
                    }

                    // Creates a file stream for the output document
                    using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        // Saves the modified document with updated checkbox states
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
