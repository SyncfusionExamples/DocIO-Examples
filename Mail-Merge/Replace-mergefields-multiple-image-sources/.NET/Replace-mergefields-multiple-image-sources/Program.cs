using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Net;

namespace ReplaceMergeFieldsWithImages
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the template Word document
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
            {
                // Attach the event handler for merging image fields
                document.MailMerge.MergeImageField += MergeField_ProductImage;

                // Define merge fields and corresponding image sources
                string[] fieldNames = { "ImageFromByteArray", "ImageFromStream", "ImageFromBase64", "ImageFromURL" };
                string[] fieldValues = { "Picture1.png", "Picture2.png", "base64.txt", "https://www.syncfusion.com/downloads/support/directtrac/general/AdventureCycle-1316159971.png" };

                // Perform the mail merge operation
                document.MailMerge.Execute(fieldNames, fieldValues);

                // Save the modified document to output
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
                {
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        /// <summary>
        /// Event handler to insert images into merge fields.
        /// </summary>
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            string tempFilePath = null;
            FileStream imageStream = null;

            if (args.FieldName == "ImageFromStream")
            {
                // Load image from file as a stream
                string filePath = Path.Combine("Data", args.FieldValue.ToString());
                imageStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            }
            else if(args.FieldName == "ImageFromByteArray")
            {
                byte[] imageBytes = File.ReadAllBytes(@"Data/" + args.FieldValue.ToString());
                tempFilePath = Path.GetTempFileName();
                File.WriteAllBytes(tempFilePath, imageBytes);
                imageStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read);
            }
            else if (args.FieldName == "ImageFromBase64")
            {
                // Convert Base64 string to byte array and save as a temp file
                string base64String = File.ReadAllText(Path.Combine("Data", args.FieldValue.ToString()));
                byte[] imageBytes = Convert.FromBase64String(base64String);
                tempFilePath = Path.GetTempFileName();
                File.WriteAllBytes(tempFilePath, imageBytes);
                imageStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read);
            }
            else if (args.FieldName == "ImageFromURL")
            {
                // Download image and save as a temp file
                WebClient client = new WebClient();
                byte[] urlImageBytes = client.DownloadData(args.FieldValue.ToString());
                tempFilePath = Path.GetTempFileName();
                File.WriteAllBytes(tempFilePath, urlImageBytes);
                imageStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read);
            }

            if (imageStream != null)
            {
                args.ImageStream = imageStream;
                WPicture picture = args.Picture;
                picture.Height = 80;
                picture.Width = 150;
            }
        }
    }
}
