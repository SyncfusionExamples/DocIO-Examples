using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Apply_Style_for_TextRange_in_Paragraph
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    WCharacterStyle customStyle = wordDocument.AddCharacterStyle("MyCustomStyle") as WCharacterStyle;

                    //Set style properties
                    customStyle.CharacterFormat.FontName = "Calibri";
                    customStyle.CharacterFormat.FontSize = 18;
                    customStyle.CharacterFormat.Bold = true;
                    customStyle.CharacterFormat.UnderlineStyle = Syncfusion.Drawing.UnderlineStyle.Single;
                    //Get text range
                    WTextRange textRange = null;
                    if (wordDocument.LastParagraph.ChildEntities.Count > 0)
                        textRange = wordDocument.LastParagraph.ChildEntities[0] as WTextRange;
                    //Apply custom style
                    if (textRange != null)
                        textRange.ApplyStyle("MyCustomStyle");
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        wordDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}