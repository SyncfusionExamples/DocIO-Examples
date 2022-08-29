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
            
                //Loads the template document
                using (WordDocument document = new WordDocument(Path.GetFullPath(@"C:\Users\ElizabethAtienoOdhia\Downloads\Input (2).docx"), FormatType.Docx))
                {
                    //Replaces the header placeholder text with desired image     
                    WParagraph headerParagraph = new WParagraph(document);
                    //Aligns the paragraph horizontally to the right
                    headerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                    WPicture picture = headerParagraph.AppendPicture(Image.FromFile(@"C:\Users\ElizabethAtienoOdhia\Downloads\AdventureCycle.jpg")) as WPicture;
                    //Sets width and height for the paragraph
                    picture.Height = 65f;
                    picture.Width = 200f;
                    //Represent the part of the textbody item in Header
                    TextBodyPart headerBodyPart = new TextBodyPart(document);
                    headerBodyPart.BodyItems.Add(headerParagraph);
                    //Replaces all entries of a given regular expression with the text body part along with its formatting in header
                    document.Replace(new Regex("^<<(.*)>>"), headerBodyPart, false);


                    WParagraph footerParagraph = new WParagraph(document);
                    //Aligns the paragraph horizontally to the right
                    footerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                    //Adds the text
                    footerParagraph.AppendText(" Page ");
                    //Adds page number field to the document
                    footerParagraph.AppendField(" CurrentPageNumber", FieldType.FieldPage);
                    // Adds the text
                    footerParagraph.AppendText(" of ");
                    //Adds number of page field to the document
                    footerParagraph.AppendField(" TotalNumberOfPages ", FieldType.FieldNumPages);
                    //Represent the part of the textbody item in Footer
                    TextBodyPart footerBodyPart = new TextBodyPart(document);
                    footerBodyPart.BodyItems.Add(footerParagraph);
                    //replaces all entries of a given regular expression with the text body part along with its formatting in footer
                    document.Replace(new Regex("^//(.*)"), footerBodyPart, false);
                    //Saves and closes the document.
                    document.Save(Path.GetFullPath("Sample.docx"), FormatType.Docx);              
                }
        }
    }
}
