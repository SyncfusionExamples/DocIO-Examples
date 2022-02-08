using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_radical_equation
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
                //Sets false to show degree in radical.
                officeMathRadical.HideDegree = false;
                //Adds a degree for radical equation.
                IOfficeMathRunElement officeMathRunElement =
                officeMathRadical.Degree.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "2";
                //Adds an equation for radical.
                officeMathRunElement =
                officeMathRadical.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets the text for radical equation.
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
