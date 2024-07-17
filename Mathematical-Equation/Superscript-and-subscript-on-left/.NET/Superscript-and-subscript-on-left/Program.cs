using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Superscript_and_subscript_on_left
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
                //Adds a left subsuperscript equation.
                IOfficeMathLeftScript officeMathLeftSubScript = officeMath.Functions.Add(0, MathFunctionType.LeftSubSuperscript) as IOfficeMathLeftScript;
                //Adds run element for left subscript.
                IOfficeMathRunElement officeMathRunElement = officeMathLeftSubScript.Subscript.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for subscript.
                (officeMathRunElement.Item as WTextRange).Text = "1";
                //Adds a run element for left superscript.
                officeMathRunElement = officeMathLeftSubScript.Superscript.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for left superscript.
                (officeMathRunElement.Item as WTextRange).Text = "n";
                officeMathRunElement = officeMathLeftSubScript.Equation.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for equation.
                (officeMathRunElement.Item as WTextRange).Text = "Y";
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
