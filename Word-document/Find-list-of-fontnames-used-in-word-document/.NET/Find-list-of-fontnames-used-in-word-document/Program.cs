using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_list_of_fontnames_used_in_word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the Word document
            WordDocument document = new WordDocument(Path.GetFullPath("Data/Template.docx"));         
            // Get all font names used in the document
            List<string> fontNames = document.FontSettings.GetUsedFontNames();
            foreach (string fontName in fontNames)
            {
                Console.WriteLine(fontName);
            }
            // Closes the Word document
            document.Close();
        }
    }
}