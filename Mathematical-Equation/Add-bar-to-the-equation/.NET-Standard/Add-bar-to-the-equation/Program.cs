using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Add_bar_to_the_equation
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
                //Adds an bar function.
                IOfficeMathBar mathBar = officeMath.Functions.Add(0, MathFunctionType.Bar) as IOfficeMathBar;
                //Sets the bar top.
                mathBar.BarTop = true;
                //Adds the run element for bar.
                IOfficeMathRunElement officeMathRunElement = mathBar.Equation.Functions.Add(0, MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                //Sets text for bar equation.
                (officeMathRunElement.Item as WTextRange).Text = "a";
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
