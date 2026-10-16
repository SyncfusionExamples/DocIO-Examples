using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.IO;
using System.Collections.Generic;
namespace Remove_previous_Paragraph_of_PlaceHolder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string documentPath = Path.GetFullPath(@"Data/Template.docx");
            // Define an array of phrases to remove
            string[] phrases = {
            "{Themes and styles also help keep your document coordinated}",
            "{Reading is easier, too, in the new Reading view}"
             };

            RemovePhrasesFromWord(documentPath, phrases);
        }
        public static void RemovePhrasesFromWord(string documentPath, string[] searchPhrases)
        {
            // Load the Word document
            using (WordDocument document = new WordDocument(documentPath, FormatType.Docx))
            {
                foreach (string searchPhrase in searchPhrases) // Iterate through all phrases
                {
                    Console.WriteLine($"Processing phrase: {searchPhrase}");
                    TextSelection[] selections = document.FindAll(searchPhrase, false, false); // Find all occurrences

                    if (selections.Length > 0)
                    {
                        foreach (var selection in selections)
                        {
                            WParagraph paragraph = selection.GetAsOneRange().OwnerParagraph;
                            if (paragraph?.PreviousSibling is WParagraph prevParagraph)
                                prevParagraph.OwnerTextBody?.ChildEntities.Remove(prevParagraph);
                            paragraph?.OwnerTextBody?.ChildEntities.Remove(paragraph);
                        }
                    }
                }
                // Save the modified document
                document.Save(Path.GetFullPath(@"../../../Output/output.docx"), FormatType.Docx);
            }
        }
    }
}
