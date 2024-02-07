using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Data;
using System.IO;

namespace Exception_field_not_in_data_source
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
                    //Sets “ClearFields” to true to remove empty mail merge fields from document.
                    document.MailMerge.ClearFields = false;
                    //Uses the mail merge event to throw exception if the field is not in the data source.
                    document.MailMerge.BeforeClearField += new BeforeClearFieldEventHandler(BeforeClearFieldEvent);
                    //Execute mail merge.
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

        #region Helper Methods
        /// <summary>
        /// Throws exception if the field is not in the data source.
        /// </summary>
        private static void BeforeClearFieldEvent(object sender, BeforeClearFieldEventArgs args)
        {
            try
            {
                if (!args.HasMappedFieldInDataSource)
                {
                    throw new Exception($"The field {args.FieldName} is not defined in data source");
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>		
        private static DataTable GetDataTable()
        {
            //Create an instance of DataTable.
            DataTable dataTable = new DataTable("Employee");
            //Add columns.
            dataTable.Columns.Add("EmployeeId");
            dataTable.Columns.Add("City");
            //dataTable.Columns.Add("Designation");
            //Add records.
            DataRow row;
            row = dataTable.NewRow();
            row["EmployeeId"] = "1001";
            row["City"] = null;
            //row["Designation"] = "SD";
            dataTable.Rows.Add(row);

            row = dataTable.NewRow();
            row["EmployeeId"] = "1002";
            row["City"] = "";
            //row["Designation"] = null;
            dataTable.Rows.Add(row);

            row = dataTable.NewRow();
            row["EmployeeId"] = "1003";
            row["City"] = "London";
            dataTable.Rows.Add(row);

            return dataTable;
        }
        #endregion
    }
}
