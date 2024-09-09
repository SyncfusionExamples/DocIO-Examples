using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Add_box_to_the_equation
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
                //Adds a box equation.
                IOfficeMathBox mathBox = officeMath.Functions.Add(0, MathFunctionType.Box) as IOfficeMathBox;
                //Adds the run element for box.
                IOfficeMathRunElement officeMathRunElement =
                officeMath.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for math.
                (officeMathRunElement.Item as WTextRange).Text = "a+b";
                //Enables the flag, to behave the box and its contents as a single operator.
                mathBox.OperatorEmulator = true;
                //Enables the flag, to act box as the mathematical differential.
                mathBox.EnableDifferential = true;
                //Adds a break in box equation.
                mathBox.Break = officeMath.Breaks.Add(0);
                //Adds the run element for box.
                officeMathRunElement =
                mathBox.Equation.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for box equation.
                (officeMathRunElement.Item as WTextRange).Text = "==";
                //Adds the run element for box.
                officeMathRunElement =
                mathBox.Equation.Functions.Add(1, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for box equation.
                (officeMathRunElement.Item as WTextRange).Text = "adx";
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
