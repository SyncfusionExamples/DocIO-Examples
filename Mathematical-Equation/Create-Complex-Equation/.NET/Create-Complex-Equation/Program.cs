using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_Complex_Equation
{
    class Program
    {
        static void Main(string[] args)
        {

            // Create WordDocument instance
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                // Add mathml equation in the word document
                AddMathMLEquation(document.LastParagraph);
                // Save the document to file
                document.Save(Path.GetFullPath(@"Output/Output.docx"));
            }
        }
        /// <summary>
        /// Adds a math equation in the word document using MathML
        /// </summary>
        /// <param name="paragraph">Represents a Word document paragraph to add math text</param>
        private static void AddMathMLEquation(WParagraph paragraph)
        {
            WordDocument document = paragraph.Document;
            //Creates a new MathML element
            WMath math = paragraph.AppendMath();
            IOfficeMath officeMath = math.MathParagraph.Maths.Add();

            // Adds a Subscript equation
            IOfficeMathScript mathScript = AddMathScript(officeMath, MathScriptType.Subscript);
            //Adds a math text
            AddMathText(document, mathScript.Equation, "L");
            AddMathText(document, mathScript.Script, "r")
                ;
            AddMathText(document, officeMath, "=");

            AddMathText(document, officeMath, "1.95");
            mathScript = AddMathScript(officeMath, MathScriptType.Subscript);
            AddMathText(document, mathScript.Equation, "f");
            AddMathText(document, mathScript.Script, "t");

            //Adds a math fraction
            IOfficeMathFraction mathFraction = officeMath.Functions.Add(MathFunctionType.Fraction) as IOfficeMathFraction;
            //Adds a numerator text
            AddMathText(document, mathFraction.Numerator, "E");
            //Adds a math script
            mathScript = AddMathScript(mathFraction.Denominator, MathScriptType.Subscript);

            //Adds a math text for Superscript
            AddMathText(document, mathScript.Equation, "F");
            AddMathText(document, mathScript.Script, "L");

            //Adds a radical equation
            IOfficeMathRadical officeMathRadical = officeMath.Functions.Add(MathFunctionType.Radical) as IOfficeMathRadical;
            //Sets false to show degree in radical
            officeMathRadical.HideDegree = true;
            //Adds a numerator text
            AddMathText(document, officeMathRadical.Degree, "");

            officeMath = officeMathRadical.Equation;

            mathFraction = officeMath.Functions.Add(MathFunctionType.Fraction) as IOfficeMathFraction;
            //Adds a numerator text
            AddMathText(document, mathFraction.Numerator, "J");
            mathScript = AddMathScript(mathFraction.Denominator, MathScriptType.Subscript);

            //Adds a math text for Superscript
            AddMathText(document, mathScript.Equation, "S");
            AddMathText(document, mathScript.Script, "xc");
            mathScript = AddMathScript(mathFraction.Denominator, MathScriptType.Subscript);

            //Adds a math text for Superscript
            AddMathText(document, mathScript.Equation, "h");
            AddMathText(document, mathScript.Script, "0");

            AddMathText(document, officeMath, "+");

            //Adds a radical equation
            officeMathRadical = officeMath.Functions.Add(MathFunctionType.Radical) as IOfficeMathRadical;
            //Sets false to show degree in radical
            officeMathRadical.HideDegree = true;
            //Adds a numerator text
            AddMathText(document, officeMathRadical.Degree, "");

            officeMath = officeMathRadical.Equation;

            //Adds a math script element
            IOfficeMathScript mathScript1 = AddMathScript(officeMath, MathScriptType.Superscript);

            IOfficeMathDelimiter mathDelimiter = mathScript1.Equation.Functions.Add(MathFunctionType.Delimiter) as IOfficeMathDelimiter;

            // Adds an office math in the delimiter
            officeMath = mathDelimiter.Equation.Add() as IOfficeMath;

            mathFraction = officeMath.Functions.Add(MathFunctionType.Fraction) as IOfficeMathFraction;
            //Adds a numerator text
            AddMathText(document, mathFraction.Numerator, "J");
            mathScript = AddMathScript(mathFraction.Denominator, MathScriptType.Subscript);

            //Adds a math text for Superscript
            AddMathText(document, mathScript.Equation, "S");
            AddMathText(document, mathScript.Script, "xc");
            mathScript = AddMathScript(mathFraction.Denominator, MathScriptType.Subscript);

            //Adds a math text for Superscript
            AddMathText(document, mathScript.Equation, "h");
            AddMathText(document, mathScript.Script, "0");
            //Adds a math text
            AddMathText(document, mathScript1.Script, "2");

            officeMath = officeMathRadical.Equation;

            AddMathText(document, officeMath, "+");
            AddMathText(document, officeMath, "6.76");

            //Adds a math script element
            mathScript1 = AddMathScript(officeMath, MathScriptType.Superscript);
            mathDelimiter = mathScript1.Equation.Functions.Add(MathFunctionType.Delimiter) as IOfficeMathDelimiter;
            // Adds an office math in the delimiter
            officeMath = mathDelimiter.Equation.Add() as IOfficeMath;
            //Adds a math fraction
            mathFraction = officeMath.Functions.Add(MathFunctionType.Fraction) as IOfficeMathFraction;
            //Adds a numerator text
            AddMathText(document, mathFraction.Denominator, "E");
            //Adds a math script
            mathScript = AddMathScript(mathFraction.Numerator, MathScriptType.Subscript);
            //Adds a math text for Superscript
            AddMathText(document, mathScript.Equation, "F");
            AddMathText(document, mathScript.Script, "L");
            //Adds a math text
            AddMathText(document, mathScript1.Script, "2");
        }
        /// <summary>
        /// Adds a math text
        /// </summary>
        /// <param name="document">Represents a Word document to add math text</param>
        /// <param name="officeMath">Represents an office math to add math text</param>
        /// <param name="text">Represents the text to set for math item</param>
        private static IOfficeMathRunElement AddMathText(WordDocument document, IOfficeMath officeMath, string text)
        {
            //Adds math text
            IOfficeMathRunElement officeMathParaItem = officeMath.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
            officeMathParaItem.Item = new WTextRange(document);
            //Set math text value
            (officeMathParaItem.Item as WTextRange).Text = text;
            return officeMathParaItem;
        }
        /// <summary>
        /// Adds a math Subscript or Superscript equation
        /// </summary>
        private static IOfficeMathScript AddMathScript(IOfficeMath officeMath, MathScriptType mathScriptType)
        {
            IOfficeMathScript mathScript = officeMath.Functions.Add(MathFunctionType.SubSuperscript) as IOfficeMathScript;
            //Sets the script type as Subscript or Superscript
            mathScript.ScriptType = mathScriptType;
            return mathScript;
        }
    }
}
