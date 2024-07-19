using System.IO;
using Syncfusion.DocIO;
using System.Reflection;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.OfficeChart;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Interactive;
using Syncfusion.Pdf.Parsing;

namespace Rename_PDF_Bookmarks_From_Word
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
                        //Sets ExportBookmarks for preserving Word document headings as PDF bookmarks.
                        renderer.Settings.ExportBookmarks = ExportBookmarkType.Bookmarks;
                        //Converts Word document into PDF document
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            //Saves the PDF document to MemoryStream.
                            using (MemoryStream stream = new MemoryStream())
                            {
                                pdfDocument.Save(stream);
                                stream.Position = 0;
                                //Load the PDF document.
                                using (PdfLoadedDocument pdfLoadedDocument = new PdfLoadedDocument(stream))
                                {
                                    int i = 1;
                                    //Get each bookmark and changes the title of the bookmark.
                                    foreach (PdfBookmark pdfBookmark in pdfLoadedDocument.Bookmarks)
                                    {
                                        pdfBookmark.Title = "PdfBookMark" + i++;
                                    }
                                    //Saves the PDF file to file system.    
                                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../WordToPDF.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                                    {
                                        pdfLoadedDocument.Save(outputStream);
                                    }
                                }
                            }
                        } 
                    }
                }
            }
        }
    }
}
