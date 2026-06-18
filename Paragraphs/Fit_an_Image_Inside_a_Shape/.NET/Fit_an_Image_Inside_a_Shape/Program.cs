using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


namespace Fit_an_Image_Inside_a_Shape
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

                    if (picture != null && picture.OwnerParagraph.OwnerTextBody.Owner is WTextBox shape)
                    {
                        float boxWidth = shape.TextBoxFormat.Width
                                          - shape.TextBoxFormat.InternalMargin.Left
                                          - shape.TextBoxFormat.InternalMargin.Right;

                        float boxHeight = shape.TextBoxFormat.Height
                                          - shape.TextBoxFormat.InternalMargin.Top
                                          - shape.TextBoxFormat.InternalMargin.Bottom;

                        picture.Width = boxWidth;
                        picture.Height = boxHeight;
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
