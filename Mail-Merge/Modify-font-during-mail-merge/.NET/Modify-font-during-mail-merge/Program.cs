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
            args.TextRange.CharacterFormat.FontName = "Arial";
            args.TextRange.CharacterFormat.FontSize = 18;
            args.TextRange.CharacterFormat.Bold = true;
            args.TextRange.CharacterFormat.Italic = true;
            args.TextRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
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
