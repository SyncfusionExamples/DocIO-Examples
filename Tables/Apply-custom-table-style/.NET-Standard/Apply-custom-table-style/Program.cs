using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_custom_table_style
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    WSection section = document.Sections[0];
                    WTable table = section.Tables[0] as WTable;
                    //Adds a new custom table style.
                    WTableStyle tableStyle = document.AddTableStyle("CustomStyle") as WTableStyle;
                    //Applies formatting for whole table.
                    tableStyle.TableProperties.RowStripe = 1;
                    tableStyle.TableProperties.ColumnStripe = 1;
                    tableStyle.TableProperties.Paddings.Top = 0;
                    tableStyle.TableProperties.Paddings.Bottom = 0;
                    tableStyle.TableProperties.Paddings.Left = 5.4f;
                    tableStyle.TableProperties.Paddings.Right = 5.4f;
                    //Applies conditional formatting for first row.
                    ConditionalFormattingStyle firstRowStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstRow);
                    firstRowStyle.CharacterFormat.Bold = true;
                    firstRowStyle.CharacterFormat.TextColor = Color.FromArgb(255, 255, 255, 255);
                    firstRowStyle.CellProperties.BackColor = Color.Blue;
                    //Applies conditional formatting for first column.
                    ConditionalFormattingStyle firstColumnStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.FirstColumn);
                    firstColumnStyle.CharacterFormat.Bold = true;
                    //Applies conditional formatting for odd row.
                    ConditionalFormattingStyle oddRowBandingStyle = tableStyle.ConditionalFormattingStyles.Add(ConditionalFormattingType.OddRowBanding);
                    oddRowBandingStyle.CellProperties.BackColor = Color.WhiteSmoke;
                    //Applies the custom table style to the table.
                    table.ApplyStyle("CustomStyle");
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
