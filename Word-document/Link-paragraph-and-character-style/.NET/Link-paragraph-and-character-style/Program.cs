using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Link_paragraph_and_character_style
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a Word document.
            using (WordDocument document = new WordDocument())
            {
                //This method adds a section and a paragraph in the document.
                document.EnsureMinimal();
                //Adds a new paragraph style named "ParagraphStyle".
                WParagraphStyle paraStyle = document.AddParagraphStyle("ParagraphStyle") as WParagraphStyle;
                //Sets the formatting of the style.
                paraStyle.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Adds a new character style named "CharacterStyle".
                IWCharacterStyle charStyle = document.AddCharacterStyle("CharacterStyle");
                //Sets the formatting of the style.
                charStyle.CharacterFormat.Bold = true;
                charStyle.CharacterFormat.Italic = true;
                //Link both paragraph and character style.
                paraStyle.LinkedStyleName = "CharacterStyle";
                //Appends the contents into the paragraph.
                document.LastParagraph.AppendText("AdventureWorks Cycles");
                //Applies the style to paragraph.
                document.LastParagraph.ApplyStyle("ParagraphStyle");
                //Appends new paragraph in section.
                document.LastSection.AddParagraph();
                //Appends the contents into the paragraph.
                document.LastParagraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Applies style to the text range.
                (document.LastParagraph.ChildEntities[0] as WTextRange).ApplyStyle("ParagraphStyle");
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
