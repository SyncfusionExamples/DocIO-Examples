using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Data;
using System.IO;

namespace Event_to_bind_data_for_unmerged_fields
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
                    //Uses the mail merge event to clear the unmerged field while perform mail merge execution.
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
		/// Represents the method that handles BeforeClearField event.
		/// </summary>
		private static void BeforeClearFieldEvent(object sender, BeforeClearFieldEventArgs args)
		{
			if (args.HasMappedFieldInDataSource)
			{
				//To check whether the mapped field has null value.
				if (args.FieldValue == null || args.FieldValue == DBNull.Value)
				{
					//Gets the unmerged field name.
					string unmergedFieldName = args.FieldName;
					string ownerGroup = args.GroupName;
					//Sets error message for unmerged fields.
					args.FieldValue = "Error! The value of MergeField " + unmergedFieldName + " of owner group " + ownerGroup + " is defined as Null in the data source.";
				}
				else
					//If field value is empty, you can set whether the unmerged merge field can be clear or not.
					args.ClearField = true;
			}
			else
			{
				string unmergedFieldName = args.FieldName;
				//Sets error message for unmerged fields, which is not found in data source.
				args.FieldValue = "Error! The value of MergeField " + unmergedFieldName + " is not found in the data source.";
			}
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
			//Add records.
			DataRow row;
			row = dataTable.NewRow();
			row["EmployeeId"] = "1001";
			row["City"] = null;
			dataTable.Rows.Add(row);

			row = dataTable.NewRow();
			row["EmployeeId"] = "1002";
			row["City"] = "";
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
