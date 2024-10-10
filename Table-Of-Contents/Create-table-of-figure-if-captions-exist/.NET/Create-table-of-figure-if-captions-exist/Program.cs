using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;

namespace Create_table_of_figure_if_captions_exist
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the existing Word document from file stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/InputDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Define the target caption name (e.g., "Figure")
                    string targetCaption = "Figure";

                    // Find all SEQ fields in the document with the given caption name
                    List<Entity> seqField = document.FindAllItemsByProperty(EntityType.SeqField, "CaptionName", targetCaption);

                    // Check if any SEQ fields with the target caption exist
                    if (seqField.Count > 0)
                    {
                        // Create a new paragraph for the "List of Figures" title
                        WParagraph paragraph = new WParagraph(document);
                        paragraph.AppendText("List of Figures");
                        // Apply Heading1 style to the paragraph
                        paragraph.ApplyStyle(BuiltinStyle.Heading1);
                        // Insert the paragraph at the beginning of the document
                        document.LastSection.Body.ChildEntities.Insert(0, paragraph);

                        // Create a new paragraph for the Table of Contents (TOC)
                        paragraph = new WParagraph(document);
                        // Append the TOC for figures (based on SEQ fields)
                        TableOfContent tableOfContent = paragraph.AppendTOC(1, 3);

                        // Exclude heading style paragraphs from TOC entries
                        tableOfContent.UseHeadingStyles = false;

                        // Set the SEQ field identifier for the table of figures (targeting "Figure")
                        tableOfContent.TableOfFiguresLabel = "Figure";

                        // Exclude caption labels and numbers in TOC entries
                        tableOfContent.IncludeCaptionLabelsAndNumbers = false;

                        // Insert the TOC paragraph into the document
                        document.LastSection.Body.ChildEntities.Insert(1, paragraph);
                    }

                    // Update the Table of Contents to reflect any changes
                    document.UpdateTableOfContents();

                    // Save the modified document to a new file
                    using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(docStream1, FormatType.Docx);
                    }
                }
            }
        }
    }
}
