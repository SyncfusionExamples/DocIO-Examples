using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Data;
using System.IO;

namespace Modify_font_during_mail_merge
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Uses the mail merge events to perform the conditional formatting during runtime.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(ModifyFont);
                    //Executes Mail Merge with groups.
                    document.MailMerge.ExecuteGroup(GetDataTable());
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
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
        private static void ModifyFont(object sender, MergeFieldEventArgs args)
        {
            // Sets the font name to Arial for the selected text range.
            args.TextRange.CharacterFormat.FontName = "Arial";
            // Sets the font size to 18 points for the selected text range.
            args.TextRange.CharacterFormat.FontSize = 18;
            // Applies bold formatting to the selected text range.
            args.TextRange.CharacterFormat.Bold = true;
            // Applies italic formatting to the selected text range.
            args.TextRange.CharacterFormat.Italic = true;
            // Applies single underline to the selected text range.
            args.TextRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
        }

        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        private static DataTable GetDataTable()
        {
            // Creates a new DataTable with the name "Employee".
            DataTable dataTable = new DataTable("Employee");
            // Adds a column "EmployeeName" to the DataTable.
            dataTable.Columns.Add("EmployeeName");
            // Adds a column "EmployeeNumber" to the DataTable.
            dataTable.Columns.Add("EmployeeNumber");

            // Loops 20 times to add rows to the DataTable.
            for (int i = 0; i < 20; i++)
            {
                // Creates a new DataRow for the DataTable.
                DataRow datarow = dataTable.NewRow();
                // Adds the newly created DataRow to the DataTable.
                dataTable.Rows.Add(datarow);
                // Sets the value of the first column (EmployeeName) for the current row.
                datarow[0] = "Employee" + i.ToString();
                // Sets the value of the second column (EmployeeNumber) for the current row.
                datarow[1] = "EMP" + i.ToString();
            }
            // Returns the populated DataTable.
            return dataTable;
        }
        #endregion
    }
}
