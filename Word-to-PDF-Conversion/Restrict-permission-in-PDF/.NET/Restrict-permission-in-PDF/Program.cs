using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Security;

namespace Restrict_permission_in_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            //Document security.
                            PdfSecurity security = pdfDocument.Security;
                            //Specifies key size and encryption algorithm using 256-bit key in AES mode.
                            security.KeySize = PdfEncryptionKeySize.Key256Bit;
                            security.Algorithm = Syncfusion.Pdf.Security.PdfEncryptionAlgorithm.AES;
                            security.OwnerPassword = "syncfusion";
                            //It restrict printing and copying of PDF document.
                            security.Permissions = ~(PdfPermissionsFlags.CopyContent | PdfPermissionsFlags.Print);
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
