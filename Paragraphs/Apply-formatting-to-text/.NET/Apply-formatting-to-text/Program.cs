using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;
using System.IO;

namespace Apply_formatting_to_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add new section to the document.
                IWSection section = document.AddSection();
                //Add new paragraph to the section.
                IWParagraph firstParagraph = section.AddParagraph();
                //Add new text to the paragraph.
                IWTextRange firstText = firstParagraph.AppendText("This is the first text range. ");
                //Apply formatting for first text range.
                firstText.CharacterFormat.Bold = true;
                firstText.CharacterFormat.FontSize = 14;
                firstText.CharacterFormat.Shadow = true;
                firstText.CharacterFormat.SmallCaps = true;
                IWTextRange secondText = firstParagraph.AppendText("This the second text range");
                //Apply formatting for second text range.
                secondText.CharacterFormat.HighlightColor = Color.GreenYellow;
                secondText.CharacterFormat.UnderlineStyle = UnderlineStyle.DotDash;
                secondText.CharacterFormat.Italic = true;
                secondText.CharacterFormat.FontName = "Times New Roman";
                secondText.CharacterFormat.TextColor = Color.Green;
                //Add new paragraph to the section.
                IWParagraph secondParagraph = section.AddParagraph();
                //Add new text to the paragraph.
                IWTextRange thirdText = secondParagraph.AppendText("שלום עולם");
                thirdText.CharacterFormat.Bidi = true;
                //Set language Identifier for right to left characters.
                thirdText.CharacterFormat.LocaleIdBidi = (short)LocaleIDs.he_IL;
                //Add third paragraph to the section.
                IWParagraph thirdParagraph = section.AddParagraph();
                //Add text to the third paragraph.
                IWTextRange fourthText = thirdParagraph.AppendText("X");
                IWTextRange fifthText = thirdParagraph.AppendText("2");
                //Apply super script formatting for fifth text range.
                fifthText.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                IWParagraph fourthParagraph = section.AddParagraph();
                //Add text to the fourth paragraph.
                IWTextRange sixthText = fourthParagraph.AppendText("m");
                IWTextRange seventhText = fourthParagraph.AppendText("3");
                //Apply sub script formatting for seventh text range.
                seventhText.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;
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
