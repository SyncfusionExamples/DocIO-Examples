using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_fraction_equation
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
                //Adds a fraction equation.
                IOfficeMathFraction mathFraction =
                officeMath.Functions.Add(0, MathFunctionType.Fraction) as IOfficeMathFraction;
                //Sets the denominator for fraction.
                IOfficeMathRunElement officeMathRunElement =
                mathFraction.Numerator.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "a";
                //Sets the numerator for fraction.
                officeMathRunElement =
                mathFraction.Denominator.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "b";
                //Sets the fraction type.
                mathFraction.FractionType = MathFractionType.NormalFractionBar;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
