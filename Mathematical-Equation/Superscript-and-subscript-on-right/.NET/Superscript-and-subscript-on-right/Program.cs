using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Superscript_and_subscript_on_right
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
                //Adds a right subsuperscript equation.
                IOfficeMathRightScript officeMathRightScript = officeMath.Functions.Add(0, MathFunctionType.RightSubSuperscript) as IOfficeMathRightScript;
                //Sets false to align subscript and superscript horizontally.
                officeMathRightScript.IsSkipAlign = true;
                //Adds run element for right subscript.
                IOfficeMathRunElement officeMathRunElement = officeMathRightScript.Subscript.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for right subscript.
                (officeMathRunElement.Item as WTextRange).Text = "1";
                //Adds run element for right superscript.
                officeMathRunElement = officeMathRightScript.Superscript.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for right superscript.
                (officeMathRunElement.Item as WTextRange).Text = "n";
                officeMathRunElement = officeMathRightScript.Equation.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for equation.
                (officeMathRunElement.Item as WTextRange).Text = "Y";
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
