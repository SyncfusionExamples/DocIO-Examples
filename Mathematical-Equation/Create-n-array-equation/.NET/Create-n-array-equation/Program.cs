using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_n_array_equation
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
                //Adds a N-Array equation.
                IOfficeMathNArray officeMathNArray = officeMath.Functions.Add(0, MathFunctionType.NArray) as IOfficeMathNArray;
                //Sets N-Array character.
                officeMathNArray.NArrayCharacter = "∑";
                //Enables the flag, to grow N-array character to full height of the arguments.
                officeMathNArray.HasGrow = false;
                //Enables the flag to hide lower limit.
                officeMathNArray.HideLowerLimit = false;
                //Enables the flag to hide upper limit.
                officeMathNArray.HideUpperLimit = false;
                //Enables the flag to set limit position as SubSuperscript.
                officeMathNArray.SubSuperscriptLimit = true;
                IOfficeMathRunElement officeMathRunElement =
                officeMathNArray.Subscript.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for superscript property of NArray equation.
                (officeMathRunElement.Item as WTextRange).Text = "n=1";
                officeMathRunElement =
                officeMathNArray.Superscript.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "10";
                officeMathRunElement =
                officeMathNArray.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for NArray equation.
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
