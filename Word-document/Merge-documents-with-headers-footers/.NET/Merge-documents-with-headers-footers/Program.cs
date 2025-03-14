using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;

namespace Merge_documents_with_headers_footers
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream destStream = new FileStream(@"Data/DestinationDocument.docx", FileMode.Open, FileAccess.Read))
            {
                using (WordDocument destinationDocument = new WordDocument(destStream, FormatType.Docx))
                {
                    //Get source document from the specified directory
                    string[] sourceData = Directory.GetFiles(@"Data/sourceDocuments");
                    foreach (string inputFile in sourceData)
                    {
                        using (FileStream sourceStream = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                        {
                            using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx))
                            {
                                foreach (WSection section in sourceDocument.Sections)
                                {
                                    WSection currentSection = section;
                                    // If the source document doesn't have a header or footer, it will retain the destination's header or footer when merged.
                                    // To prevent this, add an empty paragraph to the source's header or footer
                                    if (section.HeadersFooters.FirstPageHeader.Count == 0)
                                        section.HeadersFooters.FirstPageHeader.AddParagraph();
                                    if (section.HeadersFooters.FirstPageFooter.Count == 0)
                                        section.HeadersFooters.FirstPageFooter.AddParagraph();
                                    if (section.HeadersFooters.OddHeader.Count == 0)
                                        section.HeadersFooters.OddHeader.AddParagraph();
                                    if (section.HeadersFooters.OddFooter.Count == 0)
                                        section.HeadersFooters.OddFooter.AddParagraph();
                                    if (section.HeadersFooters.EvenHeader.Count == 0)
                                        section.HeadersFooters.EvenHeader.AddParagraph();
                                    if (section.HeadersFooters.EvenFooter.Count == 0)
                                        section.HeadersFooters.EvenFooter.AddParagraph();
                                }
                                // Import the entire content of the source document into the destination document.
                                destinationDocument.ImportContent(sourceDocument);
                            }
                        }
                    }
                    using (FileStream outputStream = new FileStream(@"Output/CombinedDocument.docx", FileMode.Create, FileAccess.ReadWrite))
                    {
                        destinationDocument.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
