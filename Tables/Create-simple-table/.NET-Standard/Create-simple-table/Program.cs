using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Create_simple_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds a section into Word document.
                IWSection section = document.AddSection();
                section.AddParagraph();
                
                //Table without data.
                IWTable table = TableWithData(section);
                //Adds a paragraph
                section.AddParagraph();
                //Table with data.
                table = TableWithoutData(section);

                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        private static WTable TableWithoutData(IWSection section)
        {
            //Adds a new table into Word document.
            IWTable table = section.AddTable();
            //Specifies the total number of rows & columns.
            table.ResetCells(5, 5);
            //Accesses the instance of the cell (first row, first cell) and adds the content into cell.
            IWTextRange textRange = table[0, 0].AddParagraph().AppendText("Product - SKU");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (first row, second cell) and adds the content into cell.
            textRange = table[0, 1].AddParagraph().AppendText("Product - Price");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (second row, first cell) and adds the content into cell.
            textRange = table[0, 2].AddParagraph().AppendText("Widget Count");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (second row, second cell) and adds the content into cell.
            textRange = table[0, 3].AddParagraph().AppendText("Extended Price");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (third row, first cell) and adds the content into cell.
            textRange = table[0, 4].AddParagraph().AppendText("Type");
            textRange.CharacterFormat.Bold = true;
            WTableCell cell = table[1, 0];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("Free Stuff");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            cell = table[1, 1];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            cell = table[1, 2];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            cell = table[1, 3];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            cell = table[1, 4];
            textRange = cell.AddParagraph().AppendText("");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 0].AddParagraph().AppendText("Free Stuff");
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 1].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 2].AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 3].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            cell = table[2, 4];
            textRange = cell.AddParagraph().AppendText("");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 0].AddParagraph().AppendText("Free Stuff");
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 1].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 2].AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 3].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            cell = table[3, 4];
            textRange = cell.AddParagraph().AppendText("");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[4, 1].AddParagraph().AppendText("TOTAL");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[4, 2].AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[4, 3].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange.CharacterFormat.Bold = true;
            return table as WTable;
        }

        private static WTable TableWithData(IWSection section)
        {
            //Adds a new table into Word document.
            IWTable table = section.AddTable();
            //Specifies the total number of rows & columns.
            table.ResetCells(5, 5);
            //Accesses the instance of the cell (first row, first cell) and adds the content into cell.
            IWTextRange textRange = table[0, 0].AddParagraph().AppendText("Product - SKU");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (first row, second cell) and adds the content into cell.
            textRange = table[0, 1].AddParagraph().AppendText("Product - Price");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (second row, first cell) and adds the content into cell.
            textRange = table[0, 2].AddParagraph().AppendText("Widget Count");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (second row, second cell) and adds the content into cell.
            textRange = table[0, 3].AddParagraph().AppendText("Extended Price");
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (third row, first cell) and adds the content into cell.
            textRange = table[0, 4].AddParagraph().AppendText("Type");
            textRange.CharacterFormat.Bold = true;
            WTableCell cell = table[1, 0];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("Free Stuff");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            cell = table[1, 1];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            cell = table[1, 2];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            cell = table[1, 3];
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = cell.AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            cell = table[1, 4];
            textRange = cell.AddParagraph().AppendText("Maintenance");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 0].AddParagraph().AppendText("Free Stuff");
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 1].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 2].AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[2, 3].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            cell = table[2, 4];
            textRange = cell.AddParagraph().AppendText("Maintenance");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 0].AddParagraph().AppendText("Free Stuff");
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 1].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 2].AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[3, 3].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            cell = table[3, 4];
            textRange = cell.AddParagraph().AppendText("Maintenance");
            cell.LastParagraph.ParagraphFormat.LineSpacing = 12;
            cell.LastParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            cell.LastParagraph.ParagraphFormat.BeforeSpacing = 0;
            cell.LastParagraph.ParagraphFormat.AfterSpacing = 0;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[4, 1].AddParagraph().AppendText("TOTAL");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[4, 2].AddParagraph().AppendText("0");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
            textRange = table[4, 3].AddParagraph().AppendText("$0.00");
            textRange.OwnerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRange.CharacterFormat.Bold = true;
            return table as WTable;
        }
    }
}
