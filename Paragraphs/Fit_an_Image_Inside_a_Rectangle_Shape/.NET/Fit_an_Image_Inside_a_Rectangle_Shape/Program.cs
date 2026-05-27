using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


namespace Fit_an_Image_Inside_a_Rectangle_Shape
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the input Word document as a file stream
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    Entity entity = document.FindItemByProperty(EntityType.Picture, "Title", "Product");
                    WPicture picture = entity as WPicture;
                    if(picture.OwnerParagraph.OwnerTextBody.Owner is WTextBox shape)
                    {
                        picture.Height = shape.TextBoxFormat.Height;
                        picture.Width = shape.TextBoxFormat.Width;
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
