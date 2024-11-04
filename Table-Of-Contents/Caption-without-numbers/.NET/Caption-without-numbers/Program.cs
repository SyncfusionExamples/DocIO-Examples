using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Caption_without_numbers
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a new section to the document.
                IWSection section = document.AddSection();
                //Create a new paragraph and append a Table of Contents (TOC).
                IWParagraph paragraph = section.AddParagraph();
                TableOfContent tableOfContent = paragraph.AppendTOC(1, 3);
                //Disable a flag to exclude heading style paragraphs in TOC entries.
                tableOfContent.UseHeadingStyles = false;
                //Set the name of SEQ field identifier for table of figures.
                tableOfContent.TableOfFiguresLabel = "Figure";

                //Add a new paragraph for the first image.
                paragraph = section.AddParagraph();
                //Add the first image to the paragraph.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/MountainCycle.jpg"), FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = paragraph.AppendPicture(imageStream);
                //Add an image caption.
                paragraph = picture.AddCaption("Figure", CaptionNumberingFormat.Number, CaptionPosition.AfterImage);
                //Add text to the caption paragraph.
                paragraph.AppendText(" " + "Mountain-Cycle");
                //Apply formatting to the caption.
                paragraph.ParagraphFormat.BeforeSpacing = 8;
                paragraph.ParagraphFormat.AfterSpacing = 8;
                //Hide the caption numbering.
                WSeqField field = paragraph.ChildEntities[1] as WSeqField;
                field.HideResult = true;
                
                //Add a new paragraph for the second image.
                paragraph = section.AddParagraph();
                //Add the second image to the paragraph.
                imageStream = new FileStream(Path.GetFullPath(@"Data/RoadCycle.jpg"), FileMode.Open, FileAccess.ReadWrite);
                picture = paragraph.AppendPicture(imageStream);
                //Add an image caption.
                paragraph = picture.AddCaption("Figure", CaptionNumberingFormat.Number, CaptionPosition.AfterImage);
                //Add text to the caption paragraph.
                paragraph.AppendText(" " + "Road-Cycle");
                //Apply formatting to the caption.
                paragraph.ParagraphFormat.BeforeSpacing = 8;
                paragraph.ParagraphFormat.AfterSpacing = 8;
                //Hide the caption numbering.
                field = paragraph.ChildEntities[1] as WSeqField;
                field.HideResult = true;

                //Update the fields in the Word document.
                document.UpdateDocumentFields();
                //Update the Table of Contents (TOC).
                document.UpdateTableOfContents();
                //Create file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
