

using Microsoft.VisualBasic.FileIO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_text_heading_paragraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    for (int headingLevel = 1; headingLevel < 10; headingLevel++)
                    {
                        //Find headings based on the levels and endnote by paragraph in Word document.
                        List<Entity> headings = document.FindAllItemsByProperty(EntityType.Paragraph, "StyleName", "Heading " + headingLevel);
                        //Replace the headings with text.
                        for (int index = 0; index < headings.Count; index++)
                        {
                            WParagraph paragraph = headings[index] as WParagraph;
                            paragraph.ChildEntities.Clear();
                            paragraph.AppendText("Replaced Heading"+headingLevel+" text");
                        }
                    }
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
}