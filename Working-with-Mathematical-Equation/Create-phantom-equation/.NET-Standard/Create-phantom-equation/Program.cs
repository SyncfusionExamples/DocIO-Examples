using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_phantom_equation
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
                WMath wmath = document.LastParagraph.AppendMath();
                IOfficeMath officeMath = wmath.MathParagraph.Maths.Add();
                //Adds a radical equation.
                IOfficeMathRadical officeMathRadical = officeMath.Functions.Add(0, MathFunctionType.Radical) as IOfficeMathRadical;
                IOfficeMathRunElement officeMathRunElement =
                officeMathRadical.Degree.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "2";
                //Adds a phantom equation in radical.
                IOfficeMathPhantom officeMathPhantom =
                officeMathRadical.Equation.Functions.Add(0, MathFunctionType.Phantom) as IOfficeMathPhantom;
                //Enables the flag, to show the contents of phantom.
                officeMathPhantom.Show = true;
                //Enables the flag, to transparent the phantom.
                officeMathPhantom.Transparent = true;
                //Enables the flag, to ignore the ascent of the phantom contents in spacing.
                officeMathPhantom.ZeroAscent = true;
                //Enables the flag, to ignore the descent of the phantom contents in spacing.
                officeMathPhantom.ZeroDescent = true;
                //Enables the flag, to ignore the width of a phantom contents in spacing
                officeMathPhantom.ZeroWidth = true;
                //Adds a run element for math phantom.
                officeMathRunElement = officeMathPhantom.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for phantom equation.
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
