using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Modify_font_during_mail_merge
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NMaF5cXmBCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdmWX1cdnRRQ2NcUkZwXUo=");
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Gets the DataTable
                    DataTable dataTable = GetDataTable();
                    //Triggers the event
                    document.MailMerge.MergeImageField += MergeField_ProductImage;
                    //Performs mail merge
                    document.MailMerge.Execute(dataTable);
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
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            if (args.FieldName == "Photo")
            {
                args.Picture.Height = 85;
                args.Picture.Width = 140;
            }
        }
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDataTable()
        {
            //Creates new DataTable instance 
            DataTable table = new DataTable();
            //Add columns in DataTable
            table.Columns.Add("Photo", typeof(byte[]));
            table.Columns.Add("ProductName");
            table.Columns.Add("ProductNumber");
            table.Columns.Add("Size");
            table.Columns.Add("Weight");
            table.Columns.Add("Price");

            //Add record in new DataRow
            DataRow row = table.NewRow();
            byte[] imageBytes = File.ReadAllBytes(@"Data/image1.gif");
            row["Photo"] = imageBytes;
            row["ProductName"] = "Mountain-200";
            row["ProductNumber"] = "BK-M68B-38";
            row["Size"] = "38";
            row["Weight"] = "25";
            row["Price"] = "$2,294.99";
            table.Rows.Add(row);

            return table;
        }
        #endregion
    }
}
