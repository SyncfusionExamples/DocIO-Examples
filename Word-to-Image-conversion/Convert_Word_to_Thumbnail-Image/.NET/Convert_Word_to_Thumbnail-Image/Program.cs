using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.PdfToImageConverter;

namespace Convert_Word_document_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the Word document file stream. 
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            // Save PDF to memory stream
                            using MemoryStream pdfStream = new MemoryStream();
                            pdfDocument.Save(pdfStream);
                            pdfStream.Position = 0;
                            //Initialize PDF image converter
                            PdfToImageConverter imageConverter = new PdfToImageConverter();
                            //Load the PDF document
                            imageConverter.Load(pdfStream);
                            //Convert first page of PDF to image (thumbnail)
                            Stream thumbnailStream = imageConverter.Convert(0, false, false);
                            //Reset stream position
                            thumbnailStream.Position = 0;
                            //Save the image
                            using (FileStream file = new FileStream(Path.GetFullPath(@"Output/Image.png"), FileMode.Create))
                            {
                                thumbnailStream.CopyTo(file);
                            }
                            thumbnailStream.Dispose();
                            imageConverter.Dispose();
                        }
                    }
                }
            }
        }
    }
}