using Syncfusion.DocIO.DLS; 
using Syncfusion.DocIO;
using System.IO; 
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf; 
using Syncfusion.Pdf.Parsing; 

namespace Convert_5_pages_in_Word_to_PDF
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the input Word document
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Create a converter to convert the Word document to PDF
                    using (DocToPDFConverter render = new DocToPDFConverter())
                    {
                        // Convert the Word document to a PDF document and save it to a MemoryStream
                        using (MemoryStream pdfStream = new MemoryStream())
                        {
                            using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                            {
                                // Save the PDF document to the MemoryStream
                                pdfDocument.Save(pdfStream);
                                pdfStream.Position = 0; // Reset the stream position to the beginning

                                // Load the PDF document from the MemoryStream
                                using (PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdfStream))
                                {
                                    // Get the total number of pages in the PDF document
                                    int totalPages = loadedDocument.Pages.Count;
                                    if (totalPages > 5)
                                    {
                                        // Remove all pages after the 5th page
                                        for (int i = totalPages - 1; i >= 5; i--)
                                        {
                                            loadedDocument.Pages.RemoveAt(i);
                                        }
                                    }

                                    // Create a file stream to save the modified PDF
                                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../Data/First-5-pages-Output.pdf"), FileMode.Create, FileAccess.Write))
                                    {
                                        // Save the modified PDF document to the output file stream
                                        loadedDocument.Save(outputFileStream);
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
    }
}
