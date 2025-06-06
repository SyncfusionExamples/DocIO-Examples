﻿using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_uniform_image_size
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Uses the mail merge events handler for image fields.
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ProductImage);
                    //Specifies the field names and field values.
                    string[] fieldNames = new string[] { "Logo" , "Picture1", "Picture2", "Picture3"};
                    string[] fieldValues = new string[] { "Logo.png", "Picture1.gif", "Picture2.gif","Picture3.gif" };
                    //Executes the mail merge with groups.
                    document.MailMerge.Execute(fieldNames, fieldValues);
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
        /// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge.
            if (args.FieldName == "Logo" || args.FieldName == "Picture1" || args.FieldName == "Picture2" || args.FieldName == "Picture3")
            {
                string ProductFileName = args.FieldValue.ToString();
                //Gets the image from file system
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/" + ProductFileName), FileMode.Open, FileAccess.Read);
                args.ImageStream = imageStream;
                //Gets the picture, to be merged for image merge field
                WPicture picture = args.Picture;
                //Resizes the picture
                picture.Height = 50;
                picture.Width = 100;
            }
        }
        #endregion
    }
}
