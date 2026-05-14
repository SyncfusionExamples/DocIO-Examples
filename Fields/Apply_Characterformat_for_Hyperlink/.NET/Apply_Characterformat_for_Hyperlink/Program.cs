using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace Apply_Characterformat_for_Hyperlink
{
    class Program
    {

        static void Main(string[] args)
        {
            //Creates a new Word document
            using (WordDocument document = new WordDocument())
            {
                //Adds one section and one paragraph to the document
                document.EnsureMinimal();
                // Appends a hyperlink to the last paragraph of the document
                string linkUri = "https://www.syncfusion.com";
                IWField field = document.LastParagraph.AppendHyperlink(linkUri, "Syncfusion", HyperlinkType.WebLink);
                // Character format for hyperlink
                bool isItalic = false;
                bool isUnderline = true;
                bool isStrikeout = false;
                bool isBold = false;
                float fontSize = 12;
                //Format hyperlink
                IEntity entity = field;
                //Iterates to sibling items until Field End
                while (entity.NextSibling != null)
                {
                    if (entity is WTextRange)
                    {
                        WTextRange textRange = entity as WTextRange;
                        //Apply character format for text ranges
                        textRange.CharacterFormat.FontName = "Verdana";
                        textRange.CharacterFormat.FontSize = fontSize;
                        textRange.CharacterFormat.TextColor = Color.Red;
                        textRange.CharacterFormat.Bold = isBold;
                        textRange.CharacterFormat.Italic = isItalic;
                        textRange.CharacterFormat.UnderlineStyle = isUnderline ? UnderlineStyle.Single : UnderlineStyle.None;
                        textRange.CharacterFormat.Strikeout = isStrikeout;
                    }
                    else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldEnd)
                        break;
                    //Gets next sibling item.
                    entity = entity.NextSibling;
                }

                //Saves the Word document to the file stream.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
