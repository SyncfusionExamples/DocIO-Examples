using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_limit_equation
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
                WMath wMath = document.LastParagraph.AppendMath();
                IOfficeMath officeMath = wMath.MathParagraph.Maths.Add();
                //Adds function to the math.
                IOfficeMathFunction officeMathFunction =
                officeMath.Functions.Add(0, MathFunctionType.Function) as IOfficeMathFunction;
                //Adds a mathematical limit equation.
                IOfficeMathLimit officeMathLimit =
                officeMathFunction.FunctionName.Functions.Add(0, MathFunctionType.Limit) as IOfficeMathLimit;
                IOfficeMathRunElement officeMathRunElement =
                officeMathLimit.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for limit equation.
                (officeMathRunElement.Item as WTextRange).Text = "lim";
                //Sets the type of the limit.
                officeMathLimit.LimitType = MathLimitType.LowerLimit;
                IOfficeMathRunElement officeMathRunElement_limit =
                officeMathLimit.Limit.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement_limit.Item = new WTextRange(document);
                //Sets the limit value.
                (officeMathRunElement_limit.Item as WTextRange).Text = "n=0";
                officeMathLimit.LimitType = MathLimitType.LowerLimit;
                officeMathRunElement =
                officeMathFunction.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for base of the specified equation.
                (officeMathRunElement.Item as WTextRange).Text = "x";
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
