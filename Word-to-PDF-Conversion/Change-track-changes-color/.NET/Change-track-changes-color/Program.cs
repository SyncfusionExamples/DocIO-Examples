using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Change_track_changes_color
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Sets revision types to preserve track changes in  Word when converting to PDF.
                    wordDocument.RevisionOptions.ShowMarkup = RevisionType.Deletions | RevisionType.Formatting | RevisionType.Insertions;
                    //Sets the color to be used for revision bars that identify document lines containing revised information.
                    wordDocument.RevisionOptions.RevisionBarsColor = RevisionColor.Blue;
                    //Sets the color to be used for inserted content Insertion.
                    wordDocument.RevisionOptions.InsertedTextColor = RevisionColor.ClassicBlue;
                    //Sets the color to be used for deleted content Deletion.
                    wordDocument.RevisionOptions.DeletedTextColor = RevisionColor.ClassicRed;
                    //Sets the color to be used for content with changes of formatting properties.
                    wordDocument.RevisionOptions.RevisedPropertiesColor = RevisionColor.DarkYellow;
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            //Saves the PDF file to file system.    
                            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
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
