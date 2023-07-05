using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using static System.Collections.Specialized.BitVector32;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Convert_Word_Document_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream("Data/Template.docx", FileMode.Open, FileAccess.Read))
            {
                //Loads an existing  Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    //Instantiation of DocIORenderer for Word to PDF conversion
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Converts Word document into PDF document
                        using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
                        {
                            //Create FileStream to save the PDF.
                            using (FileStream outputStream = new FileStream("Result.pdf", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                            {
                                //Saves the PDF.
                                pdfDocument.Save(outputStream);
                            }
                        }
                    }
                }
            }
        }
    }
}
