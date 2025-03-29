using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf.Graphics;
using System.Data;
using System.IO;

namespace Generate_QRCode_from_URL
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads an existing Word document
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Uses the mail merge events handler for image fields
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ProductImage);
                    //Gets the DataTable
                    DataTable dataTable = GetDataTable();
                    //Performs mail merge
                    document.MailMerge.Execute(dataTable);

                    // Saves the Word document file to file system    
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            document.Save(outputStream, FormatType.Docx);
                        }
                    }
                }
            }

        /// <summary>
        /// Binds the image from QR code during Mail merge process by using MergeImageFieldEventHandler.
        /// </summary>
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from QR code during mail merge
            if (args.FieldName == "Website")
            {
                //Initialize a new PdfQRBarcode instance 
                PdfQRBarcode QRCode = new PdfQRBarcode();
                //Set the XDimension and text for barcode
                QRCode.XDimension = 3;
                QRCode.Text = args.FieldValue as string;
                //Convert the QR code to image 
                Stream barcodeImage = QRCode.ToImage(new SizeF(300, 300));
                barcodeImage.Position = 0;
                //Sets QR code image as result
                args.ImageStream = barcodeImage;
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
            table.Columns.Add("Name");
            table.Columns.Add("Website");

            //Add record in new DataRow
            DataRow row = table.NewRow();
            row["Name"] = "Google";
            row["Website"] = "http://www.google.com";
            table.Rows.Add(row);
            return table;
        }
    }
    }
