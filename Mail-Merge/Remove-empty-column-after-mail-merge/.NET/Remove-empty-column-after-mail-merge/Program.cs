using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Data;
using System.IO;

namespace Remove_empty_column_after_mail_merge
{
    class Program
    {
        //Boolean to check whether the merge field has value.
        public static bool hasCostValue = false;
        //Cell index of the merge field.
        public static int cellIndex;
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing.
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document.
                Stream docStream = File.OpenRead(Path.GetFullPath(@"Data/Template.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                //Get the table.
                WTable table = GetColumnIndex(document);
                //Get the data set.
                DataSet ds = GetData();
                //Even handler to verify if a field has a valid value.
                document.MailMerge.MergeField += new MergeFieldEventHandler(MergeField_TaskCost);
                //Execute Mail Merge with groups.
                document.MailMerge.ExecuteGroup(ds.Tables["Task_CostList"]);
                if (!hasCostValue)
                {
                    //Remove the empty column.
                    RemoveColumn(table);
                }
                //Saves and closes the Word document.
                docStream = File.Create(Path.GetFullPath(@"Output/Output.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }

        #region Helper Methods
        /// <summary>
        /// Get the column index.
        /// </summary>
        private static WTable GetColumnIndex(WordDocument document)
        {
            WTable table = null;
            //Get the merge field.
            WMergeField mergeField = document.FindItemByProperty(EntityType.MergeField, "FieldName", "Cost") as WMergeField;
            if (mergeField != null)
            {
                //Check whether the merge field is present inside a table cell.
                if (mergeField.OwnerParagraph.IsInCell)
                {
                    WTableCell cell = mergeField.OwnerParagraph.OwnerTextBody as WTableCell;
                    //Get the column index.
                    cellIndex = cell.GetCellIndex();
                    table = cell.OwnerRow.Owner as WTable;
                }
            }
            return table;
        }
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        private static DataSet GetData()
        {
            // Create a DataSet.
            DataSet ds = new DataSet();
            //List of Syncfusion products name.
            string[] products = { "Task 1", "Task 2", "Task 3", "Task 4", "Task 5" };
            //Add new Tables to the data set.
            DataRow row;
            ds.Tables.Add();
            ds.Tables.Add();
            //Add fields to the Task_CostList table.
            ds.Tables[0].TableName = "Task_CostList";
            ds.Tables[0].Columns.Add("Task");
            ds.Tables[0].Columns.Add("Cost");
            int count = 0;
            //Insert values to the table row.
            foreach (string product in products)
            {
                row = ds.Tables["Task_CostList"].NewRow();
                row["Task"] = product;
                ds.Tables["Task_CostList"].Rows.Add(row);
                count++;
            }
            return ds;
        }
        /// <summary>
        /// Remove the column.
        /// </summary>
        private static void RemoveColumn(WTable table)
        {
            //Iterate through all rows.
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                //Remove the cell present in the cellIndex.
                table.Rows[i].Cells.RemoveAt(cellIndex);
            }
        }
        #endregion

        #region Event Handlers
        /// <summary>
        /// Method to handle MergeField event to verify field value and set a flag.
        /// </summary>  
        private static void MergeField_TaskCost(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldName == "Cost" && hasCostValue && args.FieldValue != null
                && args.FieldValue != DBNull.Value && args.FieldValue != string.Empty)
            {
                hasCostValue = true;
            }
        }
        #endregion
    }
}
