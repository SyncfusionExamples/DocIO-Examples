using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;

// Loads an existing Word document.
using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Creates an instance of DocIORenderer.
    using (DocIORenderer renderer = new DocIORenderer())
    {
        // Converts Word document into PDF document.
        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
        {
            // Saves the PDF file to file system.    
            pdfDocument.Save(Path.GetFullPath(@"Output/WordToPDF.pdf"));
        }
    }
}
