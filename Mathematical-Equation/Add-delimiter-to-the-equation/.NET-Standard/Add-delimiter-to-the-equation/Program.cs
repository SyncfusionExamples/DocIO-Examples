using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Add_delimiter_to_the_equation
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
                //Adds a delimiter equation.
                IOfficeMathDelimiter mathDelimiter =
                officeMath.Functions.Add(0, MathFunctionType.Delimiter) as IOfficeMathDelimiter;
                //Sets the begin character.
                mathDelimiter.BeginCharacter = "[";
                //Sets the end character.
                mathDelimiter.EndCharacter = "]";
                //Enables the flag, to grow delimiter characters to full height of the arguments.
                mathDelimiter.IsGrow = true;
                //Sets the appearance of delimiters.
                mathDelimiter.DelimiterShape = MathDelimiterShapeType.Match;
                //Adds the run element for delimiter.
                IOfficeMathRunElement officeMathRunElement =
                mathDelimiter.Equation.Add(0).Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for delimiter equation.
                (officeMathRunElement.Item as WTextRange).Text = "a+b";
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
