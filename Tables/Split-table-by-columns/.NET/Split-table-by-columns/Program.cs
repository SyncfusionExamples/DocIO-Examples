using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Split_table_by_columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Access the first table in the Word document.
                    WTable table = document.Sections[0].Tables[0] as WTable;
                    //Get the index of the table within the document's text body.
                    int index = table.OwnerTextBody.ChildEntities.IndexOf(table);
                    //Determine the column index at which the table will be split.
                    int columnIndex = (int)Math.Floor((double)table.Rows[0].Cells.Count / 2);
                    //Split the table into two parts based on the calculated column index.
                    WTable secondTable = SplitTableByColumns(table, columnIndex);
                    
                    //Insert a heading above the first part of the table.
                    WParagraph titleParaOne = new WParagraph(document);
                    titleParaOne.AppendText("Part 1 of 2").CharacterFormat.Bold = true;
                    table.OwnerTextBody.ChildEntities.Insert(index, titleParaOne);
                    //Insert a heading above the second part of the table.
                    WParagraph titleParaTwo = new WParagraph(document);
                    titleParaTwo.AppendText("Part 2 of 2").CharacterFormat.Bold = true;
                    table.OwnerTextBody.ChildEntities.Insert(index + 2, titleParaTwo);
                    
                    //Insert the second part of the table into the Word document.
                    table.OwnerTextBody.ChildEntities.Insert(index + 3, secondTable);
                    //Create a file stream to save the modified Word document.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the modified Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Split a table into two parts based on the specified column index.
        /// </summary>
        /// <param name="originalTable">The table to split.</param>
        /// <param name="columnsToSplit">The column index where the split occurs.</param>
        /// <returns>The table with columns after the split point.</returns>
        public static WTable SplitTableByColumns(WTable originalTable, int columnsToSplit)
        {
            //Clone the original table to create the second part of the table.
            WTable clonedTable = originalTable.Clone();
            //Remove cells from the cloned table to keep only the columns after the split point.
            foreach (WTableRow row in clonedTable.Rows)
            {
                //Check if the table has enough columns before attempting to remove.
                if (columnsToSplit >= 0 && columnsToSplit < row.Cells.Count)
                {
                    //Remove cells from the first part of the table (columns before the split point).
                    for (int i = columnsToSplit; i >= 0; i--)
                    {
                        row.Cells.RemoveAt(i);
                    }
                }
            }
            //Remove cells from the original table to keep only the columns before the split point.
            foreach (WTableRow row in originalTable.Rows)
            {
                //Check if the table has enough columns before attempting to remove.
                if (columnsToSplit < row.Cells.Count)
                {
                    //Remove cells from the second part of the table (columns after the split point).
                    for (int i = row.Cells.Count - 1; i > columnsToSplit; i--)
                    {
                        row.Cells.RemoveAt(i);
                    }
                }
            }
            //Return the second part of the table.
            return clonedTable;
        }
    }
}