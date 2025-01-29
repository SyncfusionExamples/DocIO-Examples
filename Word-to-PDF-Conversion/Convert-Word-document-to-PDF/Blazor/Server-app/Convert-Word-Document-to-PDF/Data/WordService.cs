using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;

namespace Convert_Word_Document_to_PDF.Data
{
    public class WordService
    {
        public MemoryStream ConvertWordtoPDF()
        {
            //Open the file as Stream
            using (FileStream sourceStreamPath = new FileStream(@"wwwroot/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    //Instantiation of DocIORenderer for Word to PDF conversion
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document
                        using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                        {
                            //Saves the Word document to MemoryStream.
                            MemoryStream stream = new MemoryStream();
                            pdfDocument.Save(stream);
                            stream.Position = 0;
                            return stream;
                        }
                    }
                }
            }                                           
        }
    }
}