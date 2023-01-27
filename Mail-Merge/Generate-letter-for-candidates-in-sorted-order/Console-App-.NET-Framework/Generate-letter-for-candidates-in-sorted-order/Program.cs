using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Generate_letter_for_candidates_in_sorted_order
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document
                document.Open(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx);
                //Gets the data view
                DataView dataView = GetDataView();
                //Performs mail merge
                document.MailMerge.Execute(dataView);
                //Saves the Word document
                document.Save(Path.GetFullPath(@"../../Sample.docx"), FormatType.Docx);
            }
        }
        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataView GetDataView()
        {
            //Creates new DataTable instance 
            DataTable table = new DataTable();
            //Add columns
            table.Columns.Add("CandidateName");
            table.Columns.Add("DateOfJoining");
            //Add records
            DataRow row = table.NewRow();
            row["CandidateName"] = "Maria Anders";
            row["DateOfJoining"] = "1/21/2020";
            table.Rows.Add(row);
            row = table.NewRow();
            row["CandidateName"] = "Ana Trujillo";
            row["DateOfJoining"] = "1/21/2020";
            table.Rows.Add(row);
            row = table.NewRow();
            row["CandidateName"] = "Howard Stark";
            row["DateOfJoining"] = "1/22/2020";
            table.Rows.Add(row);
            row = table.NewRow();
            row["CandidateName"] = "Aria Cruz";
            row["DateOfJoining"] = "1/22/2020";
            table.Rows.Add(row);

            //Creates new DataView using DataTable
            DataView dataView = new DataView(table);
            //Sort the column in ascending order based on CandidateName
            dataView.Sort = "CandidateName ASC";
            return dataView;
        }
        #endregion
    }
}
