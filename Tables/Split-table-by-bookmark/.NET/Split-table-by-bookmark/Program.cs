using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Split_table_by_bookmark
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
                    //Get all the bookmarkstart in the cell.
                    //If bookmark mentioned anywhere in table "tr" in file level, DocIO parsed inside cell only.
                    List<Entity> bookmarkStart = document.FindAllItemsByProperty(EntityType.BookmarkStart, "OwnerParagraph.IsInCell", true.ToString());
                    //Find start and end row based on bookmark start.
                    //Then split the table.
                    for (int i = 0; i < bookmarkStart.Count; i++)
                    {
                        BookmarkStart start = bookmarkStart[i] as BookmarkStart;
                        WTableRow startRow = GetOwnerRow(start);
                        WTable table = startRow.Owner as WTable;
                        WTableRow endRow;

                        #region Find start and end row to split based on bookmark start
                        //If there is any bookmark start further in table,
                        //then get previous row of bookmark start.
                        if (i < bookmarkStart.Count - 1)
                        {
                            //Get the owner row of next bookmark start.
                            WTableRow ownerRow = GetOwnerRow(bookmarkStart[i + 1] as BookmarkStart);
                            //Get the previous row.
                            endRow = ownerRow.PreviousSibling as WTableRow;
                        }
                        //If there is no further bookmark start in table, consider last row as end of splitted table.
                        else
                            endRow = (startRow.Owner as WTable).LastRow;
                        #endregion

                        #region Split table
                        //Split the table from start row to end row.
                        WTable splittedTable = SplitTable(table, startRow.GetRowIndex(), endRow.GetRowIndex());
                        //Get the owner of table.
                        WTextBody ownerBody = table.OwnerTextBody;
                        int indexToInsert = ownerBody.ChildEntities.IndexOf(table) + i * 2 + 1;
                        //Insert paragraph to differentiate two tables.
                        //As per Microsoft Word behavior, if two tables without any paragraph between them, it will be treated as one table.
                        WParagraph paragraph = new WParagraph(document);
                        ownerBody.ChildEntities.Insert(indexToInsert, paragraph);
                        //Add splitted table, next to the paragraph.
                        indexToInsert++;
                        ownerBody.ChildEntities.Insert(indexToInsert, splittedTable);
                        #endregion
                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        #region Helper methods
        /// <summary>
        /// Split a table from the bookmarks start row to the end row.
        /// </summary>
        /// <param name="table">Table to split.</param>
        /// <param name="rowStartIndex">Start row index.</param>
        /// <param name="rowEndIndex">End row index.</param>
        static WTable SplitTable(WTable table, int rowStartIndex, int rowEndIndex)
        {
            //Clone the table.
            WTable clonedTable = table.Clone();
            //Table should have atleast 1 row and 1 cell.
            clonedTable.ResetCells(1, 1);
            
            //Clone the rows from owner table to new table one by one.
            for (int i = rowStartIndex; i <= rowEndIndex; i++)
            {
                clonedTable.Rows.Add(table.Rows[i].Clone());
            }
            //Remove the first row from cloned table.
            clonedTable.Rows.RemoveAt(0);
            
            //Delete those row from owner table.
            for (int i = rowEndIndex; i >= rowStartIndex; i--)
            {
                table.Rows.RemoveAt(i);
            }
            return clonedTable;
        }
        /// <summary>
        /// Get the owner row of bookmark start.
        /// </summary>
        /// <param name="bookmarkStart">Bookmark start to get owner row.</param>
        static WTableRow GetOwnerRow(BookmarkStart bookmarkStart)
        {
            if (bookmarkStart.OwnerParagraph.IsInCell)
            {
                return (bookmarkStart.OwnerParagraph.OwnerTextBody as WTableCell).OwnerRow;
            }
            return null;
        }
        #endregion
    }
}
