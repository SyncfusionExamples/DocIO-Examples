using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_list_restart_numbering_HTML_with_text
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Load the input Word document from file stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Open the Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Load the input Word document from file stream
                    using (FileStream htmlStream = new FileStream(Path.GetFullPath(@"Data/sample.html"), FileMode.Open, FileAccess.Read))
                    {
                        // Open the Word document
                        using (WordDocument replaceDoc = new WordDocument(htmlStream, FormatType.Html))
                        {
                            //Replace the first word with HTML file content
                            document.Replace("Tag1", replaceDoc, true, true);
                            //Get the paragraph count
                            int oldcount = document.LastSection.Paragraphs.Count;
                            //Replace the second word with the HTML file content
                            document.Replace("Tag2", replaceDoc, true, true);
                            //Enable restart numbering for the HTML file first paragraph
                            document.LastSection.Paragraphs[oldcount].ListFormat.RestartNumbering = true;
                            //Iterate through the remaining paragraphs in the document
                            for (int i = oldcount + 1; i < document.LastSection.Paragraphs.Count; i++)
                            {
                                // If existing list style presents, then continue list numbering
                                if (document.LastSection.Paragraphs[i].ListFormat.CurrentListStyle != null)
                                    document.LastSection.Paragraphs[i].ListFormat.ContinueListNumbering();
                            }

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
    }
}
