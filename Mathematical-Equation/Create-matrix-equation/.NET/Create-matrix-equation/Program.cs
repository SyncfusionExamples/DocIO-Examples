using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Create_matrix_equation
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                ///Adds one section and one paragraph to the document.
                document.EnsureMinimal();
                //Appends a new mathematical equation to the paragraph.
                WMath wmath = document.LastParagraph.AppendMath();
                IOfficeMath officeMath = wmath.MathParagraph.Maths.Add();
                //Adds matrix equation.
                IOfficeMathMatrix mathMatrix = officeMath.Functions.Add(MathFunctionType.Matrix) as IOfficeMathMatrix;
                //Sets vertical alignment for matrix.
                mathMatrix.VerticalAlignment = MathVerticalAlignment.Center;
                //Sets width for matrix columns.
                mathMatrix.ColumnWidth = 1;
                //Sets column spacing rule.
                mathMatrix.ColumnSpacingRule = SpacingRule.OneAndHalf;
                //Sets column spacing value.
                mathMatrix.ColumnSpacing = 3;
                //Enables the flag to hide place holders.
                mathMatrix.HidePlaceHolders = true;
                //Sets row spacing rule.
                mathMatrix.RowSpacingRule = SpacingRule.Double;
                //Sets row spacing value.
                mathMatrix.RowSpacing = 2;

                //Adds a new column.
                mathMatrix.Columns.Add();
                //Adds a new row.
                mathMatrix.Rows.Add();
                //Sets horizontal alignment for column.
                mathMatrix.Columns[0].HorizontalAlignment = MathHorizontalAlignment.Left;

                //Gets an argument in first cell in first row.
                officeMath = mathMatrix.Rows[0].Arguments[0];
                //Sets text for argument in first cell in first row.
                IOfficeMathRunElement officeMathRunElement = officeMath.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "1";

                //Adds a new column.
                mathMatrix.Columns.Add();
                //Adds a new row.
                mathMatrix.Rows.Add();
                //Gets an argument in second cell in first row.
                officeMath = mathMatrix.Rows[0].Arguments[1];
                //Sets text for argument in second cell in first row.
                officeMathRunElement = officeMath.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "2";

                //Gets an argument in first cell in second row.
                officeMath = mathMatrix.Rows[1].Arguments[0];
                //Sets text for argument in first cell in seond row.
                officeMathRunElement = officeMath.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "3";

                //Gets an argument in second cell in second row.
                officeMath = mathMatrix.Rows[1].Arguments[1];
                //Sets text for argument in second cell in second row.
                officeMathRunElement = officeMath.Functions.Add(MathFunctionType.RunElement) as IOfficeMathRunElement;
                officeMathRunElement.Item = new WTextRange(document);
                (officeMathRunElement.Item as WTextRange).Text = "4";
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
