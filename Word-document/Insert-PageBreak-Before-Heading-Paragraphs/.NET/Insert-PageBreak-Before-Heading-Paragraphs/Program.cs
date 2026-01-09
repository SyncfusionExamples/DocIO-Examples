using Syncfusion.DocIO.DLS;

namespace Insert_PageBreak_Before_Heading_Paragraphs
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing Word document
            WordDocument document = new WordDocument(Path.GetFullPath("Data/Input.docx"));
            // Find all paragraphs with the style "Heading 1"
            List<Entity> entities = document.FindAllItemsByProperty(EntityType.Paragraph, "StyleName", "Heading 1");
            if (entities == null)
            {
                Console.WriteLine("No paragraphs with the style 'Heading 1' found.");                
            }
            else
            {
                foreach (Entity entity in entities)
                {
                    WParagraph paragraph = entity as WParagraph;
                    // Get the index of the current paragraph in its parent text body
                    int index = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);
                    // Continue only if there is a previous entity
                    if (index > 0)
                    {
                        // Get the previous entity and cast it to a paragraph
                        WParagraph previousParagraph = paragraph.OwnerTextBody.ChildEntities[index - 1] as WParagraph;
                        bool hasPageBreak = false;
                        // Check if the previous paragraph ends with a page break
                        if (previousParagraph != null && previousParagraph.ChildEntities.Count > 0)
                        {
                            // Get the last item in the previous paragraph
                            ParagraphItem lastItem = previousParagraph.ChildEntities[previousParagraph.ChildEntities.Count - 1] as ParagraphItem;
                            // Enable the boolean if the last item is a page break
                            if (lastItem is Break && (lastItem as Break).BreakType == BreakType.PageBreak)
                                hasPageBreak = true;
                        }
                        // If no page break is found, insert the page break before the current paragraph
                        if (!hasPageBreak)
                        {
                            WParagraph newPara = new WParagraph(document);
                            newPara.AppendBreak(BreakType.PageBreak);
                            // Insert the new paragraph with page break at the correct position
                            paragraph.OwnerTextBody.ChildEntities.Insert(index, newPara);
                        }
                    }
                }
            }
             // Save the Word document.   
            document.Save(Path.GetFullPath("../../../Output/Output.docx"));
            // Close the Word document.
            document.Close();
        }
    }
}
