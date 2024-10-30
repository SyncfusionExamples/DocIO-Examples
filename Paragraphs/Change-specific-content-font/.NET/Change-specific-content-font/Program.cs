using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Change_specific_content_font
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Find all the paragraphs based on specific style name.
                    List<Entity> paragraphs = document.FindAllItemsByProperty(EntityType.Paragraph, "StyleName", "List Paragraph");
                    //Iterate through each paragraph.
                    foreach (WParagraph paragraph in paragraphs)
                    {
                        //Iterate through each child items in the paragraph.
                        foreach (Entity childItem in paragraph.ChildEntities)
                        {
                            //Check if entity is WTextRange.
                            if (childItem is WTextRange)
                            {
                                WTextRange textRange = childItem as WTextRange;
                                //Change the font name for the text range.
                                textRange.CharacterFormat.FontName = "Algerian";
                            }
                        }
                    }
                    //Create output file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        //Save the modified Word document.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}