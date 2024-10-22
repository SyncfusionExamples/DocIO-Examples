using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Collections;
using System.Data;

namespace Outer_group_fields_within_inner_group
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the template Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Create the commands which contain the queries to get the data from the dataset.
                    ArrayList commands = new ArrayList();
                    commands.Add(new DictionaryEntry("Bills", ""));
                    commands.Add(new DictionaryEntry("ProductDetails", "BillId = %Bills.BillId%"));
                    commands.Add(new DictionaryEntry("Price", "DetailID = %ProductDetails.DetailID%"));
                    //Create the data set that contains data to perform the mail merge.
                    DataSet dataSet = GetDataSet();
                    //Remove groups that contain empty merge fields. 
                    document.MailMerge.RemoveEmptyGroup = true;
                    //Remove paragraphs that contain empty merge fields. 
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    //Execute the nested mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                    //Create a file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save a Word document to the file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
            #region Helpher Methods
            /// <summary>
            /// Create a data set to perform the mail merge.
            /// </summary>
            private static DataSet GetDataSet()
            {
                DataSet ds = new DataSet();
                DataTable table = new DataTable("Bills");
                table.Columns.Add("BillID");
                table.Columns.Add("ControlNumber");
                table.Columns.Add("RecipientId");
                table.Columns.Add("Picture", typeof(byte[]));
                DataRow row = table.NewRow();
                table.Rows.Add(row);
                row["BillID"] = "BL7936";
                row["ControlNumber"] = "CN100";
                row["RecipientId"] = "900893674";
                FileStream fs = new FileStream((@"Data/Mountain-300.png"), FileMode.Open, FileAccess.Read);
                byte[] buff = new byte[fs.Length];
                fs.Read(buff, 0, buff.Length);
                fs.Dispose();
                row["Picture"] = buff;
                ds.Tables.Add(table);

                table = new DataTable("ProductDetails");
                table.Columns.Add("BillID");
                table.Columns.Add("DetailID");
                table.Columns.Add("ControlNumber");
                table.Columns.Add("ProductAmount");
                table.Columns.Add("Picture", typeof(byte[]));
                row = table.NewRow();
                table.Rows.Add(row);
                row["BillID"] = "BL7936";
                row["DetailID"] = "6758671";
                row["ControlNumber"] = "CN110";
                row["ProductAmount"] = "1500";
                fs = new FileStream((@"Data/Road-550-W.png"), FileMode.Open, FileAccess.ReadWrite);
                buff = new byte[fs.Length];
                fs.Read(buff, 0, buff.Length);
                fs.Dispose();
                row["Picture"] = buff;
                ds.Tables.Add(table);

                table = new DataTable("Price");
                table.Columns.Add("DetailID");
                table.Columns.Add("ControlNumber");
                table.Columns.Add("DiscountAmount");
                table.Columns.Add("Picture", typeof(byte[]));
                row = table.NewRow();
                table.Rows.Add(row);
                row["DetailID"] = "6758671";
                row["ControlNumber"] = "CN111";
                row["DiscountAmount"] = "500";
                fs = new FileStream((@"Data/Mountain-200.png"), FileMode.Open, FileAccess.ReadWrite);
                buff = new byte[fs.Length];
                fs.Read(buff, 0, buff.Length);
                fs.Dispose();
                row["Picture"] = buff;
                ds.Tables.Add(table);
                return ds;
            }
            #endregion        
    }
}
