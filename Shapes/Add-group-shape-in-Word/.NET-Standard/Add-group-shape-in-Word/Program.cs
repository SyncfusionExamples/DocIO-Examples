using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_group_shape_in_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                WParagraph paragraph = section.AddParagraph() as WParagraph;
                //Creates new group shape.
                GroupShape groupShape = new GroupShape(document);
                //Adds group shape to the paragraph.
                paragraph.ChildEntities.Add(groupShape);
                //Creates new shape.
                Shape shape = new Shape(document, AutoShapeType.RoundedRectangle);
                //Sets height and width for shape.
                shape.Height = 100;
                shape.Width = 150;
                //Sets horizontal and vertical position.
                shape.HorizontalPosition = 72;
                shape.VerticalPosition = 72;
                //Set wrapping style for shape.
                shape.WrapFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets horizontal and vertical origin.
                shape.HorizontalOrigin = HorizontalOrigin.Page;
                shape.VerticalOrigin = VerticalOrigin.Page;
                //Adds the specified shape to group shape.
                groupShape.Add(shape);
                //Creates new picture.
                WPicture picture = new WPicture(document);
                using (FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Image.png"), FileMode.Open, FileAccess.ReadWrite))
                {
                    picture.LoadImage(imageStream);
                }
                //Sets wrapping style for picture.
                picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets height and width for the image.
                picture.Height = 100;
                picture.Width = 100;
                //Sets horizontal and vertical position.
                picture.HorizontalPosition = 400;
                picture.VerticalPosition = 150;
                //Sets horizontal and vertical origin.
                picture.HorizontalOrigin = HorizontalOrigin.Page;
                picture.VerticalOrigin = VerticalOrigin.Page;
                //Adds the specified picture to group shape.
                groupShape.Add(picture);
                //Creates new textbox.
                WTextBox textbox = new WTextBox(document);
                textbox.TextBoxFormat.Width = 150;
                textbox.TextBoxFormat.Height = 75;
                //Adds new text to the textbox body.
                IWParagraph textboxParagraph = textbox.TextBoxBody.AddParagraph();
                textboxParagraph.AppendText("Text inside text box");
                //Sets wrapping style for textbox.
                textbox.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.Behind;
                //Sets horizontal and vertical position.
                textbox.TextBoxFormat.HorizontalPosition = 200;
                textbox.TextBoxFormat.VerticalPosition = 200;
                //Sets horizontal and vertical origin.
                textbox.TextBoxFormat.VerticalOrigin = VerticalOrigin.Page;
                textbox.TextBoxFormat.HorizontalOrigin = HorizontalOrigin.Page;
                //Adds the specified textbox to group shape.
                groupShape.Add(textbox);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
