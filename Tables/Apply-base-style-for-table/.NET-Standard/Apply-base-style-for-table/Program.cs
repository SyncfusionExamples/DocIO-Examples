using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_base_style_for_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add one section and paragraph in document.
                document.EnsureMinimal();
                //Add a table to the Word document.
                WTable table = document.LastSection.AddTable() as WTable;
                table.ResetCells(3, 2);
                table[0, 0].AddParagraph().AppendText("Row 1 Cell 1");
                table[0, 1].AddParagraph().AppendText("Row 1 Cell 2");
                table[1, 0].AddParagraph().AppendText("Row 2 Cell 1");
                table[1, 1].AddParagraph().AppendText("Row 2 Cell 2");
                table[2, 0].AddParagraph().AppendText("Row 3 Cell 1");
                table[2, 1].AddParagraph().AppendText("Row 3 Cell2");
                //Add a new custom table style.
                WTableStyle tableStyle = document.AddTableStyle("CustomStyle1") as WTableStyle;
                tableStyle.TableProperties.RowStripe = 1;
                //Apply conditional formatting for first row.
                ConditionalFormattingStyle firstRowStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstRow);
                firstRowStyle.CharacterFormat.Bold = true;
                //Apply conditional formatting for odd row.
                ConditionalFormattingStyle oddRowBandingStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.OddRowBanding);
                oddRowBandingStyle.CharacterFormat.Italic = true;
                // Apply built in table style as base style for CustomStyle1.
                tableStyle.ApplyBaseStyle(BuiltinTableStyle.TableContemporary);
                //Applies the custom table style to the table
                table.ApplyStyle("CustomStyle1");

                document.LastSection.AddParagraph();
                //Create another table in the Word document.
                table = document.LastSection.AddTable() as WTable;
                table.ResetCells(3, 2);
                table[0, 0].AddParagraph().AppendText("Row 1 Cell 1");
                table[0, 1].AddParagraph().AppendText("Row 1 Cell 2");
                table[1, 0].AddParagraph().AppendText("Row 2 Cell 1");
                table[1, 1].AddParagraph().AppendText("Row 2 Cell 2");
                table[2, 0].AddParagraph().AppendText("Row 3 Cell 1");
                table[2, 1].AddParagraph().AppendText("Row 3 Cell2");

                //Adds a new custom table style.
                tableStyle = document.AddTableStyle("CustomStyle2") as WTableStyle;
                tableStyle.TableProperties.RowStripe = 1;
                //Apply conditional formatting for first row.
                firstRowStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstRow);
                firstRowStyle.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Apply conditional formatting for odd row.
                oddRowBandingStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.OddRowBanding);
                oddRowBandingStyle.CharacterFormat.TextColor = Color.Red;

                //Add a new custom table style.
                WTableStyle tableStyle2 = document.AddTableStyle("CustomStyle3") as WTableStyle;
                tableStyle2.TableProperties.RowStripe = 1;
                //Apply conditional formatting for first row.
                ConditionalFormattingStyle firstRowStyle2 = tableStyle2.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstRow);
                firstRowStyle2.CellProperties.BackColor = Color.Blue;
                //Apply conditional formatting for odd row.
                ConditionalFormattingStyle oddRowStyle2 = tableStyle2.ConditionalFormattingStyles.Add(ConditionalFormattingType.OddRowBanding);
                oddRowStyle2.CellProperties.BackColor = Color.Yellow;
                //Apply custom table style as base style for another custom table style.
                tableStyle2.ApplyBaseStyle("CustomStyle2");
                //Apply the custom table style to the table.
                table.ApplyStyle("CustomStyle3");

                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}