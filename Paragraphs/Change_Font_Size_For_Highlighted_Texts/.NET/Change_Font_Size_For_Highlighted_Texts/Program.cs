using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Change_Font_Size_For_Highlighted_Texts
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    // Finds all the text ranges in the document which have highlight color.
                    List<Entity> entities = document.FindAllItemsByProperty(EntityType.TextRange, "CharacterFormat.HighlightColor.IsEmpty", false.ToString());

                    // Iterates the text ranges.
                    foreach (Entity entity in entities)
                    {
                        // Casts the entity as WTextRange.
                        WTextRange textRange = entity as WTextRange;
                        // Get character format of the text
                        WCharacterFormat charFormat = textRange.CharacterFormat;
                        // Set text's font size larger
                        charFormat.FontSize = 14;
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
