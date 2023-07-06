using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection.Metadata;

namespace Replace_EMF_with_PNG_in_Word_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream docStream = new FileStream(@"../../../SyncfusionConvertWordToPdfIssueDoc.docx", FileMode.Open, FileAccess.Read))
            {
                //Load file stream into Word document.
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    ConvertEMFToPNG(wordDocument);
                    //Instantiation of DocIORenderer for Word to PDF conversion.
                    DocIORenderer render = new DocIORenderer();
                    //Convert Word document into PDF document.
                    using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
                    {
                        //Save the PDF file.
                        using (FileStream outputFile = new FileStream("Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            pdfDocument.Save(outputFile);
                        
                    }
                }
            }
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.pdf")
            {
                UseShellExecute = true
            };
            process.Start();
        }
        /// <summary>
        /// Replace EMF with PNG images.
        /// </summary>
        /// <param name="wordDocument"></param>
        private static void ConvertEMFToPNG(WordDocument wordDocument)
        {
            //Find all images by EntityType in Word document.
            List<Entity> images = wordDocument.FindAllItemsByProperty(EntityType.Picture, null, null);

            //Replace EMF with PNG images.
            for (int i = 0; i < images.Count; i++)
            {
                WPicture picture = images[i] as WPicture;
                //Convert EMG image format to PNG format with the same size.
                Image image = Image.FromStream(new MemoryStream(picture.ImageBytes));
                if (image.RawFormat.Equals(ImageFormat.Emf) || image.RawFormat.Equals(ImageFormat.Wmf))
                {
                    float height = picture.Height;
                    float width = picture.Width;
                    FileStream imgFile = new FileStream("Output.png", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    image.Save(imgFile, ImageFormat.Png);
                    imgFile.Dispose();
                    image.Dispose();

                    FileStream imageStream = new FileStream(@"Output.png", FileMode.Open, FileAccess.ReadWrite);
                    picture.LoadImage(imageStream);
                    picture.LockAspectRatio = false;
                    picture.Height = height;
                    picture.Width = width;
                    imageStream.Dispose();
                }

            }
        }
    }
}
