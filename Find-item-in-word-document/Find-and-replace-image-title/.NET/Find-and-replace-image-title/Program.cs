using Syncfusion.DocIO.DLS; 
using Syncfusion.DocIO; 

namespace Find_and_replace_image_title
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the Word template document as a file stream.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Open the Word document from the file stream.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    // Find a picture in the document by its Title property.
                    WPicture picture = document.FindItemByProperty(EntityType.Picture, "Title", "Adventure Works Cycle") as WPicture;

                    // If the picture is found, modify its Title property.
                    if (picture != null)
                    {
                        // Change the Title of the found picture.
                        picture.Title = "Mountain-200";
                    }

                    // Create a file stream for saving the modified document.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        // Save the modified Word document to the file stream in DOCX format.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
