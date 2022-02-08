using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_superscript_equation
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
                //Adds a subsuperscript equation.
                IOfficeMathScript officeMathScript = officeMath.Functions.Add(0, MathFunctionType.SubSuperscript) as IOfficeMathScript;
                //Sets the type of the script.
                officeMathScript.ScriptType = MathScriptType.Superscript;
                //Adds a run element for script.
                IOfficeMathRunElement officeMathRunElement =
                officeMathScript.Script.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                WTextRange textRange = officeMathRunElement.Item as WTextRange;
                //Sets text for script.
                textRange.Text = "2";
                //Adds run element for equation.
                officeMathRunElement =
                officeMathScript.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text.
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
