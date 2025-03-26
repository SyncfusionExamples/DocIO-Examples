using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Generate_Documents_for_each_record
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Get the data for mail merge.
                DataTable table = GetDataTable();
                //Iterate to the each row and generate mail merged document for each rows.
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    //Load file stream into Word document.
                    using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                    {
                        //Executes mail merge using the data row.
                        document.MailMerge.Execute(table.Rows[i]);

                        //Create a file stream.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/" + "Record_" + (i + 1) + ".docx"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Save the Word document to the file stream.
                            document.Save(outputFileStream, FormatType.Docx);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Get the data for mail merge.
        /// </summary>
        /// <returns></returns>
        static DataTable GetDataTable()
        {
            DataTable table = new DataTable();

            //Defining columns
            table.Columns.Add("Name");
            table.Columns.Add("Street");
            table.Columns.Add("City");
            table.Columns.Add("ProjectNo");

            //Set values
            DataRow row;
            row = table.NewRow();
            row["Name"] = "Andreas Waning";
            row["Street"] = "Middelwegg 2";
            row["City"] = "Vreden";
            row["ProjectNo"] = "4711";
            table.Rows.Add(row);

            row = table.NewRow();
            row["Name"] = "Mike Korf";
            row["Street"] = "teststreet";
            row["City"] = "TestCity";
            row["ProjectNo"] = "4711";
            table.Rows.Add(row);

            return table;
        }
    }
}
