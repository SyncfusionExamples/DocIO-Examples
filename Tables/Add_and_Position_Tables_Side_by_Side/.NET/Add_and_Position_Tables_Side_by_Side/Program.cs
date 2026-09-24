using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Add_and_Position_Tables_Side_by_Side
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of WordDocument class
           using ( WordDocument document = new WordDocument())
           {
                //Adds a section.
                IWSection section = document.AddSection();

                #region Left side Table
                //Adds the (first Around table) into the section.
                IWTable table1 = section.AddTable();
                //Set WrapTextAround as true to set the table as a floating table.
                table1.TableFormat.WrapTextAround = true;
                //Specifies the total number of rows & columns
                table1.ResetCells(2, 1);

                //Gets the first row first cell of the first table.
                WTableCell cell = table1[0, 0];
                //Formatting the cell
                cell.CellFormat.Borders.Bottom.BorderType = BorderStyle.Cleared;
                cell.CellFormat.BackColor = Syncfusion.Drawing.Color.Gray;
                cell.Width = 240;
                //Adds a paragraph into the cell.
                IWParagraph para = cell.AddParagraph();
                para.AppendText("PRESCRIBER");
                //Center aligns the text of the paragraph.
                para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                //Gets the second row of the first table.
                WTableRow row = table1.Rows[1];
                //Set the height of the row.
                row.Height = 120f;

                //Gets the second row first cell of the first table.
                cell = table1[1, 0];
                //Formatting the cell
                cell.Width = 240;
                //Adds a paragraph into the cell.
                cell.AddParagraph().AppendText("Dr. John, Bertha");
                //Adds a nested table into the cell (second row, first cell).
                IWTable nestTable = cell.AddTable();
                //Clears the cell borders.
                nestTable.TableFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;
                //Creates the specified number of rows and columns to nested table
                nestTable.ResetCells(6, 3);

                //Accesses the instance of the nested table cell (first row, first cell)
                WTableCell nestedCell = nestTable.Rows[0].Cells[0];
                nestedCell.Width = 56f;
                nestedCell.AddParagraph().AppendText("DEA#");
                //Accesses the instance of the nested table cell (first row, second cell)
                nestedCell = nestTable.Rows[0].Cells[1];
                nestedCell.Width = 10f;
                nestedCell.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (first row, third cell)
                nestedCell = nestTable.Rows[0].Cells[2];
                nestedCell.AddParagraph().AppendText("BD5541443");
                nestedCell.Width = 150f;


                //Accesses the instance of the nested table cell (second row, first cell)
                nestedCell = nestTable.Rows[1].Cells[0];
                nestedCell.Width = 56f;
                //Adds the content into nested cell
                nestedCell.AddParagraph().AppendText("LIC#");
                //Accesses the instance of the nested table cell (second row, second cell)
                nestedCell = nestTable.Rows[1].Cells[1];
                nestedCell.Width = 10f;
                nestedCell.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (second row, third cell)
                nestedCell = nestTable.Rows[1].Cells[2];
                nestedCell.AddParagraph().AppendText("199399");
                nestedCell.Width = 150f;

                //Accesses the instance of the nested table cell (third row, first cell)
                nestedCell = nestTable.Rows[2].Cells[0];
                nestedCell.Width = 56f;
                //Adds the content into nested cell
                nestedCell.AddParagraph().AppendText("NPI#");
                //Accesses the instance of the nested table cell (third row, second cell)
                nestedCell = nestTable.Rows[2].Cells[1];
                nestedCell.Width = 10f;
                nestedCell.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (third row, third cell)
                nestedCell = nestTable.Rows[2].Cells[2];
                nestedCell.AddParagraph().AppendText("1225021009");
                nestedCell.Width = 150f;


                //Accesses the instance of the nested table cell (fourth row, first cell)
                nestedCell = nestTable.Rows[3].Cells[0];
                nestedCell.Width = 56f;
                //Adds the content into nested cell
                nestedCell.AddParagraph().AppendText("Address");
                //Accesses the instance of the nested table cell (fourth row, second cell)
                nestedCell = nestTable.Rows[3].Cells[1];
                nestedCell.Width = 10f;
                nestedCell.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (fourth row, third cell)
                nestedCell = nestTable.Rows[3].Cells[2];
                nestedCell.AddParagraph().AppendText("212 E 106th St,New York, NY, 10029");
                nestedCell.Width = 150f;


                //Accesses the instance of the nested table cell (fifth row, first cell)
                nestedCell = nestTable.Rows[4].Cells[0];
                nestedCell.Width = 56f;
                //Adds the content into nested cell
                nestedCell.AddParagraph().AppendText("Phone");
                //Accesses the instance of the nested table cell (fifth row, second cell)
                nestedCell = nestTable.Rows[4].Cells[1];
                nestedCell.Width = 10f;
                nestedCell.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (fifth row, third cell)
                nestedCell = nestTable.Rows[4].Cells[2];
                nestedCell.AddParagraph().AppendText("(212) 360-2600");
                nestedCell.Width = 150f;


                //Accesses the instance of the nested table cell (sixth row, first cell)
                nestedCell = nestTable.Rows[5].Cells[0];
                nestedCell.Width = 56f;
                //Adds the content into nested cell
                nestedCell.AddParagraph().AppendText("Fax");
                //Accesses the instance of the nested table cell (sixth row, second cell)
                nestedCell = nestTable.Rows[5].Cells[1];
                nestedCell.Width = 10f;
                nestedCell.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (sixth row, tthird cell)
                nestedCell = nestTable.Rows[5].Cells[2];
                nestedCell.AddParagraph().AppendText("(212) 360-2616");
                nestedCell.Width = 150f;
                #endregion

                #region Right side Table
                //Second around table
                IWTable table2 = section.AddTable();
                //Sets the table format.
                table2.TableFormat.WrapTextAround = true;
                table2.TableFormat.Positioning.HorizPosition = 250f;
                //Creates the specified number of rows and columns to table.
                table2.ResetCells(2, 1);

                //Gets the second row first cell of the first table.
                cell = table2[0, 0];
                //Adds a paragraph into the cell.
                para = cell.AddParagraph();
                //Appends text to paragraph.
                para.AppendText("PATIENT");
                //Sets paragraph alignment as center. 
                para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                //Cell formats
                cell.CellFormat.BackColor = Syncfusion.Drawing.Color.Gray;
                cell.Width = 240;

                //Gets the second row of the second table.
                row = table2.Rows[1];
                //Sets the row height.
                row.Height = 120f;

                //Gets the second row first cell of the second table.
                cell = table2[1, 0];
                cell.Width = 240;
                cell.AddParagraph().AppendText("Alam, Carep");
                IWTable nestTable2 = cell.AddTable();
                nestTable2.TableFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Cleared;
                //Creates the specified number of rows and columns to nested table
                nestTable2.ResetCells(5, 3);

                //Accesses the instance of the nested table cell (first row, first cell)
                WTableCell nestedCell2 = nestTable2.Rows[0].Cells[0];
                nestedCell2.Width = 56f;
                //Adds the content into nested cell
                nestedCell2.AddParagraph().AppendText("DOB");
                //Accesses the instance of the nested table cell (first row, second cell)
                nestedCell2 = nestTable2.Rows[0].Cells[1];
                nestedCell2.Width = 10f;
                nestedCell2.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (first row, third cell)
                nestedCell2 = nestTable2.Rows[0].Cells[2];
                nestedCell2.AddParagraph().AppendText("06/02/1954");
                nestedCell2.Width = 150f;

                //Accesses the instance of the nested table cell (second row, first cell)
                nestedCell2 = nestTable2.Rows[1].Cells[0];
                nestedCell2.Width = 56f;
                //Adds the content into nested cell
                nestedCell2.AddParagraph().AppendText("Gender");
                //Accesses the instance of the nested table cell (second row, second cell)
                nestedCell2 = nestTable2.Rows[1].Cells[1];
                nestedCell2.Width = 10f;
                nestedCell2.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (second row, third cell)
                nestedCell2 = nestTable2.Rows[1].Cells[2];
                nestedCell2.AddParagraph().AppendText("M");
                nestedCell2.Width = 150f;

                //Accesses the instance of the nested table cell (third row, first cell)
                nestedCell2 = nestTable2.Rows[2].Cells[0];
                nestedCell2.Width = 56f;
                //Adds the content into nested cell
                nestedCell2.AddParagraph().AppendText("Facility");
                //Accesses the instance of the nested table cell (third row, second cell)
                nestedCell2 = nestTable2.Rows[2].Cells[1];
                nestedCell2.Width = 10f;
                nestedCell2.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (third row, third cell)
                nestedCell2 = nestTable2.Rows[2].Cells[2];
                nestedCell2.AddParagraph().AppendText("1");
                nestedCell2.Width = 150f;


                //Accesses the instance of the nested table cell (fourth row, first cell)
                nestedCell2 = nestTable2.Rows[3].Cells[0];
                nestedCell2.Width = 56f;
                //Adds the content into nested cell
                nestedCell2.AddParagraph().AppendText("Address");
                //Accesses the instance of the nested table cell (fourth row, second cell)
                nestedCell2 = nestTable2.Rows[3].Cells[1];
                nestedCell2.Width = 10f;
                nestedCell2.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (fourth row, third cell)
                nestedCell2 = nestTable2.Rows[3].Cells[2];
                nestedCell2.AddParagraph().AppendText("G7, Schenectady, NY, 12345");
                nestedCell2.Width = 150f;


                //Accesses the instance of the nested table cell (fifth row, first cell)
                nestedCell2 = nestTable2.Rows[4].Cells[0];
                nestedCell2.Width = 56f;
                //Adds the content into nested cell
                nestedCell2.AddParagraph().AppendText("Phone");
                //Accesses the instance of the nested table cell (fifth row, second cell)
                nestedCell2 = nestTable2.Rows[4].Cells[1];
                nestedCell2.Width = 10f;
                nestedCell2.AddParagraph().AppendText(":");
                //Accesses the instance of the nested table cell (fifth row, third cell)
                nestedCell2 = nestTable2.Rows[4].Cells[2];
                nestedCell2.AddParagraph().AppendText("(123) 231-2342");
                nestedCell2.Width = 150f;
                #endregion

                //Saves the Word document to MemoryStream
                using (FileStream stream = new FileStream(@"../../../Output/Result.docx", FileMode.OpenOrCreate))
                {
                    document.Save(stream, FormatType.Docx);
                }
           }
        }
    }
}
