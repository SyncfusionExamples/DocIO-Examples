using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Data;
using System.IO;

namespace Event_for_mail_merge_field
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Uses the mail merge events to perform the conditional formatting during runtime.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(ApplyAlternateRecordsTextColor);
                    //Executes Mail Merge with groups.
                    document.MailMerge.ExecuteGroup(GetDataTable());
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Represents the method that handles the MergeField event.
        /// </summary>      
        private static void ApplyAlternateRecordsTextColor(object sender, MergeFieldEventArgs args)
        {
            //Sets text color to the alternate mail merge record.
            if (args.RowIndex % 2 == 0)
            {
                args.TextRange.CharacterFormat.TextColor = Color.FromArgb(255, 102, 0);
            }
        }

        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        private static DataTable GetDataTable()
        {
            DataTable dataTable = new DataTable("Employee");
            dataTable.Columns.Add("EmployeeName");
            dataTable.Columns.Add("EmployeeNumber");

            for (int i = 0; i < 20; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Employee" + i.ToString();
                datarow[1] = "EMP" + i.ToString();
            }
            return dataTable;
        }
        #endregion
    }
}
