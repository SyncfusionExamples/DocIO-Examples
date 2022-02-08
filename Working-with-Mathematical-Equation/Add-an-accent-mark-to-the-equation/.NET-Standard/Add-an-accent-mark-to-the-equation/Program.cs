using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Add_an_accent_mark_to_the_equation
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds one section and one paragraph to the document.
                document.EnsureMinimal();
                //Appends a new mathematical equation  to the paragraph.
                WMath math = document.LastParagraph.AppendMath();
                //Adds a new math.
                IOfficeMath officeMath = math.MathParagraph.Maths.Add();
                //Adds an accent equation.
                IOfficeMathAccent mathAccent = officeMath.Functions.Add(MathFunctionType.Accent) as IOfficeMathAccent;
                //Sets the accent character.
                mathAccent.AccentCharacter = "̆";
                //Adds the run element for accent.
                IOfficeMathRunElement officeMathRunElement = mathAccent.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                WTextRange textRange = officeMathRunElement.Item as WTextRange;
                //Sets text for accent equation.
                textRange.Text = "a";
                //Applies character formatting for text range.
                textRange.CharacterFormat.Bold = true;
                textRange.CharacterFormat.Italic = true;
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
