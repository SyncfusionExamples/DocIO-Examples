using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_image_caption
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds a new section to the document.
                IWSection section = document.AddSection();
                //Sets margin of the section.
                section.PageSetup.Margins.All = 72;
                //Adds a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Adds image to  the paragraph.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Google.png"), FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = paragraph.AppendPicture(imageStream);
                //Adds Image caption.
                IWParagraph lastParagragh = picture.AddCaption("Figure", CaptionNumberingFormat.Roman, CaptionPosition.AfterImage);
                //Aligns the caption.
                lastParagragh.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Sets after spacing.
                lastParagragh.ParagraphFormat.AfterSpacing = 12f;
                //Sets before spacing.
                lastParagragh.ParagraphFormat.BeforeSpacing = 1.5f;
                //Adds a paragraph to the section.
                paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Adds image to  the paragraph.
                imageStream = new FileStream(Path.GetFullPath(@"Data/Yahoo.png"), FileMode.Open, FileAccess.ReadWrite);
                picture = paragraph.AppendPicture(imageStream);
                //Adds Image caption.
                lastParagragh = picture.AddCaption("Figure", CaptionNumberingFormat.Roman, CaptionPosition.AfterImage);
                //Aligns the caption.
                lastParagragh.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Sets before spacing.
                lastParagragh.ParagraphFormat.BeforeSpacing = 1.5f;
                //Updates the fields in Word document.
                document.UpdateDocumentFields();
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
