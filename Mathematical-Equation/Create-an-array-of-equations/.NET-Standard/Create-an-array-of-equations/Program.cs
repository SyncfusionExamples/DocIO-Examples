using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_an_array_of_equations
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
                //Adds an equation array.
                IOfficeMathEquationArray mathEquationArray =
                officeMath.Functions.Add(0, MathFunctionType.EquationArray) as IOfficeMathEquationArray;
                //Sets the vertical alignment for equation array.
                mathEquationArray.VerticalAlignment = MathVerticalAlignment.Center;
                //Enables the flag, to distribute the equation array equally within the container.
                mathEquationArray.ExpandEquationContainer = true;
                //Enables the flag, to expand the equations in an equation array to the maximum width.
                mathEquationArray.ExpandEquationContent = true;
                //Sets the row spacing rule.
                mathEquationArray.RowSpacingRule = SpacingRule.Multiple;
                //Adds the run element for equation array.
                IOfficeMathRunElement officeMathRunElement =
                mathEquationArray.Equation.Add(0).Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for equation array.
                (officeMathRunElement.Item as WTextRange).Text = "x+y+z=0";
                //Adds the run element for equation array.
                officeMathRunElement =
                mathEquationArray.Equation.Add(1).Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for equation array.
                (officeMathRunElement.Item as WTextRange).Text = "x+y-z=1";
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
