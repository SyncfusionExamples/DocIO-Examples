using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Replace_text_with_data_from_database
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Set datasource and database values.
                string datasource = "";
                string database = "";
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
                {
                    SqlConnection SqlConn = new SqlConnection("Data Source=" + datasource + ";Initial Catalog=" + database + ";Integrated Security=True");
                    //Retrive data from the table 'FindReplace' using SqlCommand.
                    SqlCommand sqlCommand = new SqlCommand("Select * from FindReplace");
                    sqlCommand.Connection = SqlConn;

                    //Load the data into DataTable using SqlDataAdapter.
                    SqlDataAdapter da = new SqlDataAdapter(sqlCommand);
                    da.SelectCommand.CommandTimeout = 0;
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    //Find and replace text with other text from SQL server.
                    foreach (DataRow row in dt.Rows)
                    {
                        //First column contains text to find.
                        //Second column contains replacement text.
                        document.Replace(row[dt.Columns[0]] as string, row[dt.Columns[1]] as string, false, false);
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
