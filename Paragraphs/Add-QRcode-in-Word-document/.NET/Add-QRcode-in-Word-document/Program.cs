using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf.Graphics;
using System.IO;

namespace Add_QRcode_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument())
            {
                // Inserts one section and one paragraph in the Word document
                document.EnsureMinimal();

                // Gets the last paragraph of the document
                WParagraph paragraph = document.LastParagraph;

                // Appending text to the Paragraph
                paragraph.AppendText("QR Code \n");

                // Generate QR code and get the FileStream
                FileStream qrImageStream = CreateQRCode();

                // Append QR code image to paragraph
                paragraph.AppendPicture(qrImageStream);

                // Close the stream after appending the picture
                qrImageStream.Close();

                // Saves the Word document file to file system    
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }

        static FileStream CreateQRCode()
        {
            PdfQRBarcode qrBarcode = new PdfQRBarcode();

            // Sets the Input mode to Binary mode
            qrBarcode.InputMode = InputMode.BinaryMode;

            // Automatically select the Version
            qrBarcode.Version = QRCodeVersion.Auto;

            // Set the Error correction level to high
            qrBarcode.ErrorCorrectionLevel = PdfErrorCorrectionLevel.High;

            // Set dimension for each block
            qrBarcode.XDimension = 4;
            qrBarcode.Text = "This is a sample QR Code";

            // Generate a temporary file path
            string tempFilePath = Path.GetTempFileName() + ".png";

            // Save the QR code as an image to a file
            using (Stream imageStream = qrBarcode.ToImage(new SizeF(300, 300)))
            {
                using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }
            }
            // Return FileStream for the saved image
            return new FileStream(tempFilePath, FileMode.Open, FileAccess.Read);
        }
    }
}
