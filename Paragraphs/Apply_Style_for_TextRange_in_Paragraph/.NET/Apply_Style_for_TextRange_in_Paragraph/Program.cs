using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Apply_Style_for_TextRange_in_Paragraph
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../../Data/Template.docx")))
            {
                WCharacterStyle customStyle = document.AddCharacterStyle("MyCustomStyle") as WCharacterStyle;

                //Set style properties
                customStyle.CharacterFormat.FontName = "Calibri";
                customStyle.CharacterFormat.FontSize = 18;
                customStyle.CharacterFormat.Bold = true;
                customStyle.CharacterFormat.UnderlineStyle = Syncfusion.Drawing.UnderlineStyle.Single;
                //Get text range
                WTextRange textRange = null;
                if (document.LastParagraph.ChildEntities.Count > 0)
                    textRange = document.LastParagraph.ChildEntities[0] as WTextRange;
                //Apply custom style
                if (textRange != null)
                    textRange.ApplyStyle("MyCustomStyle");

                //Saves the Word document.
                document.Save(Path.GetFullPath(@"../../../Output/Result.docx"), FormatType.Docx);
            }
        }
    }
}