using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_nested_group_shapes
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
                //Appends new shape to the document.
                Shape shape = new Shape(document, AutoShapeType.RoundedRectangle);
                //Sets height and width for shape.
                shape.Height = 100;
                shape.Width = 150;
                //Sets Wrapping style for shape.
                shape.WrapFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets horizontal and vertical position for shape.
                shape.HorizontalPosition = 72;
                shape.VerticalPosition = 72;
                //Sets horizontal and vertical origin for shape.
                shape.HorizontalOrigin = HorizontalOrigin.Page;
                shape.VerticalOrigin = VerticalOrigin.Page;
                //Adds the specified shape to group shape.
                groupShape.Add(shape);
                //Appends new picture to the document.
                WPicture picture = new WPicture(document);
                //Loads image from the file.
                using (FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.ReadWrite))
                {
                    picture.LoadImage(imageStream);
                }
                //Sets wrapping style for picture.
                picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                //Sets height and width for the picture.
                picture.Height = 100;
                picture.Width = 100;
                //Sets horizontal and vertical position for the picture.
                picture.HorizontalPosition = 400;
                picture.VerticalPosition = 150;
                //Sets horizontal and vertical origin for the picture.
                picture.HorizontalOrigin = HorizontalOrigin.Page;
                picture.VerticalOrigin = VerticalOrigin.Page;
                //Adds specified picture to the group shape.
                groupShape.Add(picture);
                //Creates new nested group shape.
                GroupShape nestedGroupShape = new GroupShape(document);
                //Appends new textbox to the document.
                WTextBox textbox = new WTextBox(document);
                //Sets width and height for the textbox.
                textbox.TextBoxFormat.Width = 150;
                textbox.TextBoxFormat.Height = 75;
                //Adds new text to the textbox body.
                IWParagraph textboxParagraph = textbox.TextBoxBody.AddParagraph();
                //Adds new text to the textbox paragraph.
                textboxParagraph.AppendText("Text inside text box");
                //Sets wrapping style for the textbox. 
                textbox.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.Behind;
                //Sets horizontal and vertical position for the textbox.
                textbox.TextBoxFormat.HorizontalPosition = 200;
                textbox.TextBoxFormat.VerticalPosition = 200;
                //Sets horizontal and vertical origin for the textbox.
                textbox.TextBoxFormat.VerticalOrigin = VerticalOrigin.Page;
                textbox.TextBoxFormat.HorizontalOrigin = HorizontalOrigin.Page;
                //Adds specified textbox to the nested group shape.
                nestedGroupShape.Add(textbox);
                //Appends new shape to the document.
                shape = new Shape(document, AutoShapeType.Oval);
                //Sets height and width for the new shape.
                shape.Height = 100;
                shape.Width = 150;
                //Sets horizontal and vertical position for the shape.
                shape.HorizontalPosition = 200;
                shape.VerticalPosition = 72;
                //Sets horizontal and vertical origin for the shape.
                shape.HorizontalOrigin = HorizontalOrigin.Page;
                shape.VerticalOrigin = VerticalOrigin.Page;
                //Sets horizontal and vertical position for the nested group shape.
                nestedGroupShape.HorizontalPosition = 72;
                nestedGroupShape.VerticalPosition = 72;
                //Adds specified shape to the nested group shape.
                nestedGroupShape.Add(shape);
                //Adds nested group shape to the group shape of the paragraph.
                groupShape.Add(nestedGroupShape);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
