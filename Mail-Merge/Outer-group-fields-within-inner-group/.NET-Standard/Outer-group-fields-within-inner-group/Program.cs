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
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                   //Creates the commands which contains the queries to get the data from dataset.
                   ArrayList commands = new ArrayList();
                   commands.Add(new DictionaryEntry("Bills", ""));
                   commands.Add(new DictionaryEntry("ProductDetails", "BillId = %Bills.BillId%"));
                   commands.Add(new DictionaryEntry("Price", "DetailID = %ProductDetails.DetailID%"));

                   //Creates the Data set that contains data to perform mail merge.
                   DataSet dataSet = GetDataSet();
                   //Removes group which contain empty merge fields.
                   document.MailMerge.RemoveEmptyGroup = true;
                   //Removes paragraphs which contain empty merge fields.
                   document.MailMerge.RemoveEmptyParagraphs = true;
                   //Excutes the nested mail merge.
                   document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
            #region Helpher Methods
            /// <summary>
            /// Creates data set
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
                FileStream fs = new FileStream((@"../../../Data/MetroStudio1.png"), FileMode.Open, FileAccess.Read);
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
                fs = new FileStream((@"../../../Data/MetroStudio2.png"), FileMode.Open, FileAccess.ReadWrite);
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
                fs = new FileStream((@"../../../Data/MetroStudio3.png"), FileMode.Open, FileAccess.ReadWrite);
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
