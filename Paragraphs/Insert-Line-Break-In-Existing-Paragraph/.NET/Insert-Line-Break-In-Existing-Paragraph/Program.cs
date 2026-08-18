using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_Line_Break_In_Existing_Paragraph
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads the template document.
            using (WordDocument document = new WordDocument(Path.GetFullPath("Data/Template.docx")))
            {
                //Gets the text body of first section.
                WTextBody textBody = document.Sections[0].Body;
                //Gets the paragraph at index 1.
                WParagraph paragraph = textBody.Paragraphs[1];
                //Creates a new instance of the line break.
                Break lineBreak = new Break(document, BreakType.LineBreak);
                //Inserts line break to the paragraph at a specific location (index).
                paragraph.ChildEntities.Insert(3, lineBreak);
                document.Save(Path.GetFullPath(@"Output/Output.docx"));
            }
        }
    }
}

