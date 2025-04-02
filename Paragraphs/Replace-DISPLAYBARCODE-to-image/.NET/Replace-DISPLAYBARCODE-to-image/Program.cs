using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf.Graphics;
using System.Text.RegularExpressions;
using SizeF = Syncfusion.Drawing.SizeF;
using System.Collections.Generic;
using System.IO;
using System;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Replace_DISPLAYBARCODE_to_image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the Word document from a file stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    //Replace specific barcode fields in the document with generated barcode images
                    ReplaceFieldwithImage(document);

                    //Save the modified document
                    using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream1, FormatType.Docx);
                    }
                    //Create a DocIORenderer instance to convert the Word document to a PDF
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Convert the Word document to a PDF
                        using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                        {
                            //Save the generated PDF to a file
                            using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                pdfDocument.Save(outputStream1);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Replaces fields with barcode images based on specific field codes in the document.
        /// </summary>
        /// <param name="document">The Word document object</param>
        private static void ReplaceFieldwithImage(WordDocument document)
        {
            // Find all fields in the document
            List<Entity> fields = document.FindAllItemsByProperty(EntityType.Field, "FieldType", "FieldUnknown");

            // Iterate over all found fields
            foreach (WField field in fields)
            {
                if (field != null)
                {
                    // Get the owner paragraph of the field
                    WParagraph ownerParagraph = field.OwnerParagraph as WParagraph;
                    // Get the index of the field within the paragraph
                    int index = ownerParagraph.ChildEntities.IndexOf(field);
                    // Split the field code to identify the type of barcode
                    string[] components = field.FieldCode.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    // If the field is a QR code (check for "displaybarcode" and "qr")
                    if (components[0].ToLower() == "displaybarcode" && components[2].ToLower() == "qr")
                    {
                        // Get the text to encode into the QR barcode
                        string qrBarcodeText = components[1];

                        // Initialize default values for optional parameters
                        int qValue = -1;
                        float sValue = -1;
                        float size = 90;

                        // Extract \q (error correction level) value
                        var qMatch = Regex.Match(field.FieldCode, @"\\q (\d+)");
                        if (qMatch.Success)
                        {
                            qValue = int.Parse(qMatch.Groups[1].Value);
                        }

                        // Extract \s (size) value
                        var sMatch = Regex.Match(field.FieldCode, @"\\s (\d+)");
                        if (sMatch.Success)
                        {
                            sValue = int.Parse(sMatch.Groups[1].Value);
                            size = sValue * 1.05f;  // Scale the size slightly
                        }

                        // Generate the QR barcode image as a byte array
                        byte[] qrCode = GenerateQRBarcodeImage(qrBarcodeText, sValue, qValue, size);

                        // Create a new picture object to hold the barcode image
                        WPicture picture = new WPicture(document);
                        picture.LoadImage(qrCode);

                        // Replace the original field with the picture (QR code)
                        ownerParagraph.ChildEntities.Remove(field);
                        ownerParagraph.ChildEntities.Insert(index, picture);
                    }
                    // If the field is a Code39 barcode (check for "displaybarcode" and "code39")
                    else if (components[0].ToLower() == "displaybarcode" && components[2].ToLower() == "code39")
                    {
                        // Get the text to encode into the Code39 barcode
                        string qrBarcodeText = components[1];

                        // Initialize flags for optional parameters
                        bool addText = false;
                        bool addCharacter = false;

                        // Check for the \t option (whether to display text below the barcode)
                        var tabMatch = Regex.IsMatch(field.FieldCode, @"\\t");
                        if (tabMatch)
                        {
                            addText = true;
                        }

                        // Check for the \d option (whether to include start/stop characters)
                        var digitMatch = Regex.IsMatch(field.FieldCode, @"\\d");
                        if (digitMatch)
                        {
                            addCharacter = true;
                        }

                        // Generate the Code39 barcode image as a byte array
                        byte[] qrCode = GenerateCODE39Image(qrBarcodeText, addCharacter, addText);

                        // Create a new picture object to hold the barcode image
                        WPicture picture = new WPicture(document);
                        picture.LoadImage(qrCode);

                        // Replace the original field with the picture (Code39 barcode)
                        ownerParagraph.ChildEntities.Remove(field);
                        ownerParagraph.ChildEntities.Insert(index, picture);
                    }
                }
            }
        }

        /// <summary>
        /// Generates a QR barcode image and converts it to a byte array.
        /// </summary>
        /// <param name="qrBarcodeText">The text to be encoded in the QR code</param>
        /// <param name="sSwitchValue">The size value (\s option) for the QR code</param>
        /// <param name="qSwitchValue">The error correction level (\q option) for the QR code</param>
        /// <param name="size">The size of the QR code image</param>
        /// <returns>A byte array representing the QR code image</returns>
        private static byte[] GenerateQRBarcodeImage(string qrBarcodeText, float sSwitchValue, int qSwitchValue, float size)
        {
            // Create a new QR barcode instance
            PdfQRBarcode qrBarCode = new PdfQRBarcode();

            // Set the text to be encoded
            qrBarCode.Text = qrBarcodeText;

            // Set the size if provided
            if (sSwitchValue != -1)
            {
                qrBarCode.XDimension = sSwitchValue;
            }

            // Set the error correction level based on the \q switch value
            if (qSwitchValue != -1)
            {
                switch (qSwitchValue)
                {
                    case 0:
                        qrBarCode.ErrorCorrectionLevel = PdfErrorCorrectionLevel.Low;
                        break;
                    case 1:
                        qrBarCode.ErrorCorrectionLevel = PdfErrorCorrectionLevel.Medium;
                        break;
                    case 2:
                        qrBarCode.ErrorCorrectionLevel = PdfErrorCorrectionLevel.Quartile;
                        break;
                    case 3:
                        qrBarCode.ErrorCorrectionLevel = PdfErrorCorrectionLevel.High;
                        break;
                }
            }

            // Generate the QR code image and return it as a byte array
            Stream barcodeImage = qrBarCode.ToImage(new SizeF(size, size));
            byte[] byteArray;
            using (MemoryStream ms = new MemoryStream())
            {
                barcodeImage.CopyTo(ms);
                byteArray = ms.ToArray();
            }
            return byteArray;
        }

        /// <summary>
        /// Generates a Code39 barcode image and converts it to a byte array.
        /// </summary>
        /// <param name="qrBarcodeText">The text to be encoded in the Code39 barcode</param>
        /// <param name="dSwitch">Whether to include start/stop characters (\d option)</param>
        /// <param name="tSwitch">Whether to display the text below the barcode (\t option)</param>
        /// <returns>A byte array representing the Code39 barcode image</returns>
        private static byte[] GenerateCODE39Image(string qrBarcodeText, bool dSwitch, bool tSwitch)
        {
            // Create a new Code39 barcode instance
            PdfCode39Barcode barcode = new PdfCode39Barcode();

            // Configure the barcode based on the provided options
            if (!tSwitch)
            {
                // If \t is not specified, don't display text
                barcode.Text = Regex.Replace(qrBarcodeText.ToUpper(), @"[^A-Z0-9\-\.\ \$\/\+\%]", "");
                barcode.TextDisplayLocation = TextLocation.None;
            }
            else
            {
                // If \t is specified, display the barcode text below the image
                barcode.Text = Regex.Replace(qrBarcodeText.ToUpper(), @"[^A-Z0-9\-\.\ \$\/\+\%]", "");
                PdfStandardFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 15);
                barcode.Font = font;
            }

            // Generate the barcode image and return it as a byte array
            Stream barcodeImage = barcode.ToImage(new SizeF(40, 40));
            byte[] byteArray;
            using (MemoryStream ms = new MemoryStream())
            {
                barcodeImage.CopyTo(ms);
                byteArray = ms.ToArray();
            }
            return byteArray;
        }
    }
}
