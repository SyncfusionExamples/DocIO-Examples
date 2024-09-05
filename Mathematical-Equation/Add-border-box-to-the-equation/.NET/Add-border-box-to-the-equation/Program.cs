using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Add_border_box_to_the_equation
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
                //Adds a border box equation.
                IOfficeMathBorderBox mathBorderBox =
                officeMath.Functions.Add(0, MathFunctionType.BorderBox) as IOfficeMathBorderBox;
                //Sets the diagonal strikethrough from lower left to upper right.
                mathBorderBox.StrikeDiagonalUp = true;
                //Sets the diagonal strikethrough from upper left to lower right.
                mathBorderBox.StrikeDiagonalDown = true;
                //Sets the horizontal strikethrough.
                mathBorderBox.StrikeHorizontal = true;
                //Sets the vertical strikethrough.
                mathBorderBox.StrikeVertical = true;
                //Enable the flag, to hide the bottom border of an equation.
                mathBorderBox.HideBottom = true;
                //Enable the flag, to hide the left border of an equation.
                mathBorderBox.HideLeft = true;
                //Enable the flag, to hide the right border of an equation.
                mathBorderBox.HideRight = false;
                //Enable the flag, to hide the top border of an equation.
                mathBorderBox.HideTop = false;
                //Adds the run element for border box.
                IOfficeMathRunElement officeMathRunElement = mathBorderBox.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for border box equation.
                (officeMathRunElement.Item as WTextRange).Text = "a+b-c";
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
