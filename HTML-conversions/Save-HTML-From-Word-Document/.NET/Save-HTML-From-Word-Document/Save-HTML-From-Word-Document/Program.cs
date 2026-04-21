using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Save_HTML_From_Word_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    int i = 0;
                    foreach (WSection section in document.Sections)
                    {
                        if (section.PageSetup.DifferentFirstPage)
                        {
                            GenerateHTML(section.HeadersFooters.FirstPageHeader, "FirstPageHeader_" + i + ".html");
                            GenerateHTML(section.HeadersFooters.FirstPageFooter, "FirstPageFooter_" + i + ".html");
                        }
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
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/TextBody.html"), FileMode.Create))
                    {
                        document.Save(outputStream, FormatType.Html);
                    }
                }                    
            }               
        }
        private static void GenerateHTML(WTextBody textBody, string outputFile)
        {
            string outputPath = Path.GetFullPath(@"../../../Output/");
            if (textBody.ChildEntities.Count > 0)
            {
                WordDocument document = new WordDocument();
                document.AddSection();
                foreach (Entity entity in textBody.ChildEntities)
                    document.LastSection.Body.ChildEntities.Add(entity.Clone());

                document.Save(outputPath + outputFile, FormatType.Html);
            }
        }

    }
}
        