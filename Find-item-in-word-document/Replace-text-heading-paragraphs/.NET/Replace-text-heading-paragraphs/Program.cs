using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_text_heading_paragraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens the input Word document.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    for (int headingLevel = 1; headingLevel < 10; headingLevel++)
                    {
                        //Get all heading paragraphs based on the heading style level.
                        List<Entity> headings = document.FindAllItemsByProperty(EntityType.Paragraph, "StyleName", "Heading " + headingLevel);
                        //Iterate through all headings in the list.
                        for (int index = 0; index < headings.Count; index++)
                        {
                            //Cast the current heading to WParagraph.
                            WParagraph paragraph = headings[index] as WParagraph;
                            //Remove all child elements from the paragraph.
                            paragraph.ChildEntities.Clear();
                            //Add new text to replace the heading content.
                            paragraph.AppendText("Replaced Heading" + headingLevel + " text");
                        }
                    }
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}