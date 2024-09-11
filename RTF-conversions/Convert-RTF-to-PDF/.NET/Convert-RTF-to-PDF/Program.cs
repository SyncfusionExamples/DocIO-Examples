using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Reflection.Metadata;

using (FileStream docStream = new FileStream(@"Data/Input.rtf", FileMode.Open, FileAccess.Read))
{
    //Loads file stream into Word document
    using (WordDocument wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic))
    {
        //Instantiation of DocIORenderer for Word to PDF conversion
        DocIORenderer render = new DocIORenderer();
        //render.Settings.AutoDetectComplexScript = true;
        //Converts Word document into PDF document
        PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
        //Releases all resources used by the Word document and DocIO Renderer objects
        render.Dispose();
        wordDocument.Dispose();
        //Saves the PDF file
        using (FileStream outputStream = new FileStream(@"Data/Output.pdf", FileMode.Create, FileAccess.Write))
        {
            pdfDocument.Save(outputStream);
        }
        //Closes the Word document
        wordDocument.Close();
        //Closes the instance of PDF document object
        pdfDocument.Close();
    }
}