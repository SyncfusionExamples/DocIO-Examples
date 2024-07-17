using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Modify_an_existing_equation
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Access the paragraph from Word document.
                    WParagraph paragraph = document.LastSection.Body.ChildEntities[0] as WParagraph;
                    //Access the mathematical equation from the paragraph.
                    WMath math = paragraph.ChildEntities[0] as WMath;
                    //Access the radical equation.
                    IOfficeMathRadical mathRadical = math.MathParagraph.Maths[0].Functions[1] as IOfficeMathRadical;
                    //Access the fraction equation in radical.
                    IOfficeMathFraction mathFraction = mathRadical.Equation.Functions[0] as IOfficeMathFraction;
                    //Access the n-array equation in fraction.
                    IOfficeMathNArray mathNAry = mathFraction.Numerator.Functions[0] as IOfficeMathNArray;
                    //Access the math script in n-array.
                    IOfficeMathScript mathScript = mathNAry.Equation.Functions[0] as IOfficeMathScript;
                    //Access the delimiter in math script.
                    IOfficeMathDelimiter mathDelimiter = mathScript.Equation.Functions[0] as IOfficeMathDelimiter;
                    //Removes the delimiter.
                    mathScript.Equation.Functions.Remove(mathDelimiter);
                    //Modifies the run element in math script.
                    IOfficeMathRunElement MathParagraphItem = mathScript.Equation.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                    MathParagraphItem.Item = new WTextRange(document);
                    //Sets the text value.
                    (MathParagraphItem.Item as WTextRange).Text = "x";
                    //Applies character format to the text.
                    (MathParagraphItem.Item as WTextRange).CharacterFormat.Italic = true;
                    (MathParagraphItem.Item as WTextRange).CharacterFormat.FontSize = 20;
                    //Applies math format to the text.
                    MathParagraphItem.MathFormat.Style = MathStyleType.Italic;
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
}
