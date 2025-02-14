using Syncfusion.DocIO.DLS;  
using Syncfusion.DocIO;      

namespace Change_font_size   
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document from the file system for reading.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Initialize the WordDocument object with the input stream.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
                {
                    // Find all text ranges (instances of text) in the document
                    List<Entity> textRanges = document.FindAllItemsByProperty(EntityType.TextRange, null, null);

                    // Loop through each text range and change the font size.
                    for (int i = 0; i < textRanges.Count; i++)
                    {
                        WTextRange textRange = textRanges[i] as WTextRange;
                        // Set font size to 15.
                        textRange.CharacterFormat.FontSize = 15;  
                    }

                    // Create an output stream to save the modified document.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        // Save the modified document in DOCX format.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
