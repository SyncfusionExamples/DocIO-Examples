using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Net;

namespace Mail_merge_using_image_from_URL
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
                    //Uses the mail merge events handler for image fields.
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ProductImage);
                    //Specifies the field names and field values.
                    string[] fieldNames = new string[] { "Logo" };
                    string[] fieldValues = new string[] { "https://www.syncfusion.com/downloads/support/directtrac/general/AdventureCycle-1316159971.png" };
                    //Executes the mail merge.
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from URL during mail merge.
            if (args.FieldName == "Logo")
            {
                string ProductFileName = args.FieldValue.ToString();
                WebClient client = new WebClient();
                //Download the image from URL as byte array.
                byte[] imageBytes = client.DownloadData(ProductFileName);
                MemoryStream ms = new MemoryStream(imageBytes);
                //Set the retrieved image from the memory stream.
                args.ImageStream = ms;

                //Gets the picture, to be merged for image merge field
                WPicture picture = args.Picture;
                //Resizes the picture.
                picture.Height = 50;
                picture.Width = 100;
            }
        }
        #endregion
    }
}
