using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;


namespace Fit_an_Image_Inside_a_Rectangle_Shape
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document
            using (WordDocument document = new WordDocument())
            {
                //Add a new section
                WSection section = document.AddSection() as WSection;
                //Add a new paragraph to the section.
                WParagraph paragraph = section.AddParagraph() as WParagraph;
                //Add a new rectangle shape
                Shape rectangle = paragraph.AppendShape(AutoShapeType.Rectangle, 150, 100);
                //Format the rectangle shape
                rectangle.VerticalPosition = 72;
                rectangle.HorizontalPosition = 72;
                //Add a new paragraph to a rectangle shape
                WParagraph para = rectangle.TextBody.AddParagraph() as WParagraph;
                //Append the picture to the paragraph
                WPicture picture = para.AppendPicture(File.ReadAllBytes("../../../Data/Mountain-200.jpg")) as WPicture;
                //Resize the picture according to rectangle shape
                picture.Width = rectangle.Width;
                picture.Height = rectangle.Height;
                picture.VerticalPosition = rectangle.VerticalPosition;
                picture.HorizontalPosition = rectangle.HorizontalPosition;
                //document.Settings.ResizeImageToFitInContainer = true;
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
