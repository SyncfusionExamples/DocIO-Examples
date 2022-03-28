using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Copy_fonts_to_linux_containers
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream docStream = new FileStream(@"Adventure.docx", FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document.
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    //Instantiation of DocIORenderer for Word to PDF conversion.
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
                        //Saves the PDF files.
                        using (FileStream outputStream = new FileStream("Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                        {
                            pdfDocument.Save(outputStream);
                        }
                    }
                }
            }
        }
    }
}
