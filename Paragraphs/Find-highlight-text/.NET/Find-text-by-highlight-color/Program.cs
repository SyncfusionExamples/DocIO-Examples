using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_an_text_by_highlight_color
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
                        // Create or overwrite the text file
                        using (StreamWriter writer = new StreamWriter(Path.GetFullPath(@"Output/result.txt"), false))
                        {
                            if (textRanges != null)
                            {
                                //Iterate and write the highlight color name of the current text range to the file
                                foreach (Entity entity in textRanges)
                                {
                                    WTextRange textRange = entity as WTextRange;
                                    writer.WriteLine($"HighlightColor: {textRange.CharacterFormat.HighlightColor.Name}");
                                    writer.WriteLine($"Text: {textRange.Text}");
                                    writer.WriteLine(); // Blank line between entries
                                }
                            }
                            else
                            {
                                // If no highlighted text ranges were found, write a message to the file
                                writer.WriteLine("No text ranges with highlight were found.");
                            }
                        }
                    }
                }
            }
        }

    }
}

