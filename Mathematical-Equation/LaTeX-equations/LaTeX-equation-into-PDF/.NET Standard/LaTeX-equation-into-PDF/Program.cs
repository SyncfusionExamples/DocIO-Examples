using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

// Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add one section and one paragraph to the document.
    document.EnsureMinimal();

    //Append an accent equation using LaTeX.
    document.LastParagraph.AppendMath(@"f\left(x\right)={a}_{0}+\sum_{n=1}^{\infty}{\left({a}_{n}\cos{\frac{n\pi{x}}{L}}+{b}_{n}\sin{\frac{n\pi{x}}{L}}\right)}");
    
    //Instantiation of DocIORenderer for Word to PDF conversion
    DocIORenderer render = new DocIORenderer();
    //render.Settings.AutoDetectComplexScript = true;
    //Converts Word document into PDF document
    PdfDocument pdfDocument = render.ConvertToPDF(document);
    //Releases all resources used by the Word document and DocIO Renderer objects
    render.Dispose();
    document.Dispose();
    //Save the Word document.
    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
    {
        pdfDocument.Save(outputStream);
    }
    //Close the Word document
    document.Close();
    //Closes the instance of PDF document object
    pdfDocument.Close();
}