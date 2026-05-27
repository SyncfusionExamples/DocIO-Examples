using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Save_HTML_From_Word_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the input Word document as a file stream
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    int i = 0;
                    // Iterate through each section in the document
                    foreach (WSection section in document.Sections)
                    {
                        // Handle first page header / footer if enabled
                        if (section.PageSetup.DifferentFirstPage)
                        {
                            GenerateHTML(section.HeadersFooters.FirstPageHeader, "FirstPageHeader_" + i + ".html");
                            GenerateHTML(section.HeadersFooters.FirstPageFooter, "FirstPageFooter_" + i + ".html");
                        }
                        // Handle even page header / footer if enabled
                        else if (section.PageSetup.DifferentOddAndEvenPages)
                        {
                            GenerateHTML(section.HeadersFooters.EvenHeader, "EvenHeader_" + i + ".html");
                            GenerateHTML(section.HeadersFooters.EvenFooter, "EvenFooter_" + i + ".html");

                        }
                        //This is the default header and footer
                        GenerateHTML(section.HeadersFooters.OddHeader, "OddHeader_" + i + ".html");
                        GenerateHTML(section.HeadersFooters.OddFooter, "OddFooter_" + i + ".html");

                        //After generating headers and footers, clear it
                        section.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                        section.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                        section.HeadersFooters.EvenHeader.ChildEntities.Clear();
                        section.HeadersFooters.EvenFooter.ChildEntities.Clear();
                        section.HeadersFooters.OddHeader.ChildEntities.Clear();
                        section.HeadersFooters.OddFooter.ChildEntities.Clear();

                        i++;
                    }
                    // Save the remaining document body content as HTM
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/TextBody.html"), FileMode.Create))
                    {
                        document.Save(outputStream, FormatType.Html);
                    }
                }                    
            }
        }
        /// </summary>
        // Generates an HTML file from the given text body(header or footer).
        /// </summary>
        /// <param name="textBody">The text body (header/footer) to convert.</param>
        /// <param name="outputFile">The output HTML file name.</param>

        private static void GenerateHTML(WTextBody textBody, string outputFile)
        {
            string outputPath = Path.GetFullPath(@"../../../Output/");
            // Check if the text body contains any content
            if (textBody.ChildEntities.Count > 0)
            {
                // Create a new Word document to hold extracted content
                WordDocument document = new WordDocument();
                document.AddSection();
                // Clone and add each entity from the source text body
                foreach (Entity entity in textBody.ChildEntities)
                    document.LastSection.Body.ChildEntities.Add(entity.Clone());

                //Save the extracted content as an HTML file
                document.Save(outputPath + outputFile, FormatType.Html);
            }
        }
    }
}
        