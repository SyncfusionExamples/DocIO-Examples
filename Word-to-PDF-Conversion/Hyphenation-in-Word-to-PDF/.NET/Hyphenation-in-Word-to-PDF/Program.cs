using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Hyphenation_in_Word_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Reads the language dictionary for hyphenation.
                        using (FileStream dictionaryStream = new FileStream(Path.GetFullPath(@"../../../Data/hyph_de_CH.dic"), FileMode.Open))
                        {
                            //Adds the hyphenation dictionary of the specified language.
                            Hyphenator.Dictionaries.Add("de-CH", dictionaryStream);
                            //Converts Word document into PDF document.
                            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                            {
                                //Saves the PDF file to file system.    
                                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../WordToPDF.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                                {
                                    pdfDocument.Save(outputStream);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
