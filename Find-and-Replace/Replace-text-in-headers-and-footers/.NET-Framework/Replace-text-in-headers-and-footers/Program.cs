using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_in_headers_and_footers
{
    class Program
    {
        static void Main(string[] args)
        {

            //Load the template document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Docx))
            {
                //Create paragraph for header.  
                WParagraph headerParagraph = new WParagraph(document);
                //Align paragraph horizontally to the right.
                headerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //Append picture in the paragraph.
                WPicture picture = headerParagraph.AppendPicture(Image.FromFile(@"../../Data/AdventureCycle.jpg")) as WPicture;
                //Set width and height for the picture.
                picture.Height = 65f;
                picture.Width = 200f;
                //Create text body part.
                TextBodyPart headerBodyPart = new TextBodyPart(document);
                headerBodyPart.BodyItems.Add(headerParagraph);
                //Replace all entries of a given regular expression with the text body part along with its formatting in header.
                document.Replace(new Regex("^<<(.*)>>"), headerBodyPart, false);

                //Create paragraph for footer.
                WParagraph footerParagraph = new WParagraph(document);
                //Align the paragraph horizontally to the right.
                footerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                //Add the text.
                footerParagraph.AppendText(" Page ");
                //Add page number field.
                footerParagraph.AppendField(" CurrentPageNumber", FieldType.FieldPage);
                //Add the text.
                footerParagraph.AppendText(" of ");
                //Add number of page field.
                footerParagraph.AppendField(" TotalNumberOfPages ", FieldType.FieldNumPages);
                //Create text body part.
                TextBodyPart footerBodyPart = new TextBodyPart(document);
                footerBodyPart.BodyItems.Add(footerParagraph);
                //Replace all entries of a given regular expression with the text body part along with its formatting in footer.
                document.Replace(new Regex("^//(.*)"), footerBodyPart, false);
                //Save the document.
                document.Save(Path.GetFullPath("../../Sample.docx"), FormatType.Docx);
            }
        }
    }
}
