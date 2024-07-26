using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Insert_as_new_row
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
					//Creates a data table.
					DataTable table = new DataTable("CompatibleVersions");
					table.Columns.Add("WordVersion");
					//Creates a new data row.
					DataRow row = table.NewRow();
					row["WordVersion"] = "Microsoft Word 97-2003";
					table.Rows.Add(row);
					row = table.NewRow();

					row["WordVersion"] = "Microsoft Word 2007";
					table.Rows.Add(row);
					row = table.NewRow();

					row["WordVersion"] = "Microsoft Word 2010";
					table.Rows.Add(row);
					row = table.NewRow();

					row["WordVersion"] = "Microsoft Word 2013";
					table.Rows.Add(row);
					row = table.NewRow();

					row["WordVersion"] = "Microsoft Word 2019";
					table.Rows.Add(row);

					//Enable the flag to insert a new row for every group in a table.
					document.MailMerge.InsertAsNewRow = true;
					//Execute mail merge.
					document.MailMerge.ExecuteGroup(table);
					//Creates file stream.
					using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
