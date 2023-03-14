using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Keep_group_of_rows_from_split_between_pages
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the Word document.
                IWSection section = document.AddSection();
                //Set number of table rows.
                int rowCount = 200;
                //Set number of row set.
                int rowSet = 3;
                //Create table with specified number of rows.
                IWTable innerTable = CreateTable(rowCount, rowSet, document);
                //Create outer table.
                IWTable outerTable = section.AddTable();
                //Keep a group of rows in same page when one of the group row placed on next page.
                KeepGroupOfRows(innerTable, outerTable, rowSet);
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
        /// <summary>
        /// Create the main table.
        /// </summary>        
        private static IWTable CreateTable(int rowCount, int rowSet, WordDocument document)
        {
            //Create inner table.
            WTable table = new WTable(document);
            //Specifie the total number of rows & columns.
            table.ResetCells(rowCount, 1);
            int dataNum = 1;
            //Add text to table's row.
            for (int tableRowIndex = 0; tableRowIndex < table.Rows.Count; tableRowIndex++)
            {
                table[tableRowIndex, 0].AddParagraph().AppendText("Data" + dataNum);
                dataNum++;
                if ((tableRowIndex + 1) % rowSet == 0)
                {
                    dataNum = 1;
                }
            }
            return table;
        }
        /// <summary>
        /// Add rows to outer table by keeping them in groups.
        /// </summary>
        private static void KeepGroupOfRows(IWTable innerTable, IWTable outerTable, int rowSet)
        {
            int innerTableRowIndex = 0;
            //Create number of tables based on row set and add it to outer table rows.
            IWTable table = outerTable.AddRow().AddCell().AddTable();
            while (innerTable.Rows.Count > 0)
            {
                table.Rows.Add(innerTable.Rows[0]);
                if ((innerTableRowIndex + 1) % rowSet == 0)
                {
                    table = outerTable.AddRow().Cells[0].AddTable();
                }
                innerTableRowIndex++;
            }
            //Format the outer table.
            outerTable.TableFormat.Borders.BorderType = BorderStyle.None;
            outerTable.TableFormat.IsBreakAcrossPages = false;
            outerTable.TableFormat.Paddings.Left = 0;
            outerTable.TableFormat.Paddings.Right = 0;
        }
    }
}
