using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Compression.Zip;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.IO.Compression;
using ZipArchive = Syncfusion.Compression.Zip.ZipArchive;

namespace WordToImageConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing Word document from the specified path.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Data/Template.docx"), FormatType.Docx))
            {
                // Initialize the ChartToImageConverter for converting charts during Word to image conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                // Set the scaling mode for charts (Normal mode reduces the file size).
                wordDocument.ChartToImageConverter.ScalingMode = ScalingMode.Normal;

                // Convert the entire Word document to images.
                Image[] images = wordDocument.RenderAsImages(ImageType.Bitmap);

                // Initialize a ZipArchive to store the converted images.
                using (ZipArchive zipArchive = new ZipArchive())
                {
                    // Set the compression level to the fastest.
                    zipArchive.DefaultCompressionLevel = (Syncfusion.Compression.CompressionLevel)CompressionLevel.Fastest;

                    int i = 0; // Initialize an index for naming the images.
                    foreach (Image image in images)
                    {
                        // Create a memory stream to hold the image data.
                        MemoryStream imageStream = new MemoryStream();

                        // Save the image to the memory stream in JPEG format.
                        image.Save(imageStream, ImageFormat.Jpeg);
                        imageStream.Position = 0;

                        // Create a ZipArchiveItem with the image stream.
                        string imageName = $"WordToImage_{i}.jpeg";
                        ZipArchiveItem item = new ZipArchiveItem(zipArchive, imageName, imageStream, true, FileAttributes.Normal);
                        zipArchive.AddItem(item);

                        i++; // Increment the index for the next image.
                    }

                    // Save the ZipArchive to a memory stream.
                    using (MemoryStream zipStream = new MemoryStream())
                    {
                        zipArchive.Save(zipStream, false);
                        zipStream.Position = 0;

                        // Write the zipStream to a file.
                        File.WriteAllBytes(Path.GetFullPath(@"../../Data/Images.zip"), zipStream.ToArray());
                    }
                }
            }
        }
    }
}
