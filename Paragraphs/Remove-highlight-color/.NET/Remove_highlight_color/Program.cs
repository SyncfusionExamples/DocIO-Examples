using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace Remove_highlight_color
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    // Find all text ranges in the document where the highlight color is Yellow
                    List<Entity> textRanges = document.FindAllItemsByProperty(EntityType.TextRange, "CharacterFormat.HighlightColor.Name", "Yellow");
                    // Check if any matching text ranges were found
                    if (textRanges != null)
                    {
                        //Iterate text ranges
                        foreach (Entity entity in textRanges)
                        {
                            WTextRange textRange = entity as WTextRange;
                            // Clear the highlight color on this text range.
                            textRange.CharacterFormat.HighlightColor = Color.Empty;
                        }
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

    }
}


