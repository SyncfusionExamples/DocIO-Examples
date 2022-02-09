using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Equation_with_grouping_character
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
                //Appends a new mathematical equation to the paragraph.
                WMath math = document.LastParagraph.AppendMath();
                //Adds a new math.
                IOfficeMath officeMath = math.MathParagraph.Maths.Add();
                //Adds a group character equation.
                IOfficeMathGroupCharacter officeMathGroupCharacter =
                officeMath.Functions.Add(0, MathFunctionType.GroupCharacter) as IOfficeMathGroupCharacter;
                //Sets the group character.
                officeMathGroupCharacter.GroupCharacter = "⏞";
                //Enables the flag to align group character at top.
                officeMathGroupCharacter.HasAlignTop = true;
                //Enables the flag to align the text and group character.
                officeMathGroupCharacter.HasCharacterTop = true;
                //Adds the run element for group character.
                IOfficeMathRunElement officeMathRunElement =
                officeMathGroupCharacter.Equation.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for group character equation.
                (officeMathRunElement.Item as WTextRange).Text = "a-b";
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
