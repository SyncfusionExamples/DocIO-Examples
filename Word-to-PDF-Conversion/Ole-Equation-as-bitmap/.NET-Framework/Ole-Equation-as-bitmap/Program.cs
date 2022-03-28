using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.IO;

namespace Ole_Equation_as_bitmap
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Creates an instance of the DocToPDFConverter.
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    //Sets a value indicating whether to preserve the Ole Equation as bitmap image while converting a Word document to PDF.
                    converter.Settings.PreserveOleEquationAsBitmap = true;
                    //Converts Word document into PDF document.
                    using (PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument))
                    {
                        //Saves the PDF file to file system.
                        pdfDocument.Save(Path.GetFullPath(@"../../WordToPDF.pdf"));
                    }
                }
            }
        }
    }
}
