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
                    customStyle.CharacterFormat.FontSize = 22;
                    customStyle.CharacterFormat.UnderlineStyle = Syncfusion.Drawing.UnderlineStyle.Single;
					
                    // Find the first occurrence of the target text
					TextSelection selection = document.Find("Adventure Works Cycles", true,true);

					if (selection != null)
					{
						// Convert selection into a single text range and apply the style.
						WTextRange textRange = selection.GetAsOneRange();
						textRange.ApplyStyle("MyCustomStyle");
					}
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