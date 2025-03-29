using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_different_document_format_as_OLE_object
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // Add a new section to the document.
                IWSection section = document.AddSection();

                // Add different types of documents as OLE objects.
                AddOleObject(document, section, "Data/Template.pdf", "Data/pdf.png", OleObjectType.AdobeAcrobatDocument);
                AddOleObject(document, section, "Data/Template.xlsx", "Data/excel.png", OleObjectType.ExcelWorksheet);
                AddOleObject(document, section, "Data/Adventure.docx", "Data/word.png", OleObjectType.WordDocument);
                AddOleObject(document, section, "Data/Sample.pptx", "Data/powerpoint.png", OleObjectType.PowerPointPresentation);

                // Save the Word document to a file.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }

        /// <summary>
        /// Adds an OLE object to the Word document.
        /// </summary>
        private static void AddOleObject(WordDocument document, IWSection section, string filePath, string imagePath, OleObjectType oleType)
        {
            // Add a new paragraph to the section.
            IWParagraph paragraph = section.AddParagraph();

            // Open the file to be embedded as an OLE object.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(filePath), FileMode.Open))
            using (FileStream imageStream = new FileStream(Path.GetFullPath(imagePath), FileMode.Open, FileAccess.ReadWrite))
            {
                // Load the image as a representation of the OLE object.
                WPicture picture = new WPicture(document);
                picture.LoadImage(imageStream);

                // Append the OLE object to the paragraph.
                WOleObject oleObject = paragraph.AppendOleObject(fileStream, picture, oleType);
                paragraph.AppendText("\n");

                // Set the display size of the OLE object.
                oleObject.OlePicture.Height = 80;
                oleObject.OlePicture.Width = 100;
            }
        }
    }
}
