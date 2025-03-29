using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.EJ2.PdfViewer;
using SkiaSharp;

namespace Replace_bookmark_content_with_OLE
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document template file in read/write mode
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                // Load the Word document into a Syncfusion DocIO instance
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Open the PDF file that will be inserted as an OLE object
                    FileStream pdfFileStream = new FileStream(Path.GetFullPath(@"Data/Adventure.pdf"), FileMode.Open, FileAccess.Read);

                    // Extract the first page of the PDF as an image (to use as a preview)
                    byte[] extractedImages = GetPDFFirstPageasImage(pdfFileStream);

                    // Create a picture instance to hold the extracted image
                    WPicture picture = new WPicture(document);
                    picture.LoadImage(extractedImages);

                    // Create a bookmark navigator to locate the target bookmark in the document
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);

                    // Move to the bookmark named "OLEObject" where the PDF will be inserted
                    bookmarkNavigator.MoveToBookmark("OLEObject");

                    // Get the content within the bookmark and clear it
                    TextBodyPart textBodyPart = bookmarkNavigator.GetBookmarkContent();
                    textBodyPart.BodyItems.Clear();

                    // Create a new paragraph to hold the OLE object
                    WParagraph paragraph = new WParagraph(document);
                    textBodyPart.BodyItems.Add(paragraph);

                    // Reopen the PDF file stream (because it was read earlier and may be at the end)
                    pdfFileStream = new FileStream(Path.GetFullPath(@"Data/Adventure.pdf"), FileMode.Open, FileAccess.Read);

                    // Insert the PDF file as an OLE object within the paragraph
                    WOleObject oleObject = paragraph.AppendOleObject(pdfFileStream, picture, OleObjectType.AdobeAcrobatDocument);

                    // Dispose of the PDF file stream after use to free resources
                    pdfFileStream.Dispose();

                    // Set the display size of the OLE object in the document
                    oleObject.OlePicture.Height = 200;
                    oleObject.OlePicture.Width = 200;

                    // Replace the bookmark's content with the new text body part containing the OLE object
                    bookmarkNavigator.ReplaceBookmarkContent(textBodyPart);

                    // Create a file stream to save the modified document
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        // Save the modified Word document to the output file
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Extracts the first page of a PDF as an image (PNG format).
        /// </summary>
        /// <param name="pdfFileStream">The file stream of the PDF.</param>
        /// <returns>A byte array containing the image data.</returns>
        private static byte[] GetPDFFirstPageasImage(FileStream pdfFileStream)
        {
            using (PdfRenderer pdfRenderer = new PdfRenderer())
            {
                // Load the PDF file into the renderer
                pdfRenderer.Load(pdfFileStream);

                // Export the first page of the PDF as an image
                using (SKBitmap bitmapimage = pdfRenderer.ExportAsImage(0))
                using (SKImage image = SKImage.FromBitmap(bitmapimage))
                using (SKData imageData = image.Encode(SKEncodedImageFormat.Png, 100))
                {
                    // Convert the image to a byte array and return it
                    return imageData.ToArray();
                }
            }
        }
    }
}
