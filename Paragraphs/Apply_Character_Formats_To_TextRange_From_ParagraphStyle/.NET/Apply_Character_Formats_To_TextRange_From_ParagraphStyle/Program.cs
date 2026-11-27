using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Apply_Character_Formats_To_TextRange_From_ParagraphStyle
{
    class Program
    {
        static void Main(string[] args)
        {

            //Opens an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath("Data/Template.docx")))
            {
                //Adds a paragraph to last section. 
                WParagraph paragraph = document.LastSection.AddParagraph() as WParagraph;
                //Appends the text.
                WTextRange textRange = paragraph.AppendText("The company continues to expand its market reach through innovation, efficient manufacturing processes, and a strong global distribution network.") as WTextRange;
                //Get the style in the Word document.
                WParagraphStyle paragraphStyle = document.Styles.FindByName("PalabraParagraph") as WParagraphStyle;
                //Apply character formats from paragraph style.
                textRange.ApplyCharacterFormat(paragraphStyle.CharacterFormat);
                //Saves and closes the Word document.
                document.Save(Path.GetFullPath("Output/Result.docx"));
            }
        }
    }
}
