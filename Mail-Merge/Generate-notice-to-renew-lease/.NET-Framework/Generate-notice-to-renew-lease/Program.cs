using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Generate_notice_to_renew_lease
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
                //Gets the DataTable
                DataTable dataTable = GetDataTable();
                //Performs mail merge
                document.MailMerge.Execute(dataTable);
                //Saves the Word document
                document.Save(Path.GetFullPath(@"../../Sample.docx"), FormatType.Docx);
            }
        }
        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDataTable()
        {
            //Creates new DataTable instance 
            DataTable table = new DataTable();
            //Add columns in DataTable
            table.Columns.Add("NoticeDate");
            table.Columns.Add("LandlordName");
            table.Columns.Add("Address");
            table.Columns.Add("City");
            table.Columns.Add("Region");
            table.Columns.Add("PostalCode");
            table.Columns.Add("Country");
            table.Columns.Add("AgreementDate");
            table.Columns.Add("StartDate");
            table.Columns.Add("EndDate");
            table.Columns.Add("AmountPerAnnum");
            table.Columns.Add("AmountPerMonth");

            //Add record in new DataRow
            DataRow row = table.NewRow();
            row["NoticeDate"] = "June 10, 2019";
            row["LandlordName"] = "Thomas Hardy";
            row["Address"] = "120 Hanover Sq.";
            row["City"] = "London";
            row["PostalCode"] = "WA1 1DP";
            row["Country"]= "UK";
            row["AgreementDate"] = "July, 10, 2016";
            row["StartDate"] = "July 10, 2019";
            row["EndDate"] = "July 10, 2020";
            row["AmountPerAnnum"] = "72000";
            row["AmountPerMonth"] = "6000";
            table.Rows.Add(row);

            return table;
        }
        #endregion
    }
}
