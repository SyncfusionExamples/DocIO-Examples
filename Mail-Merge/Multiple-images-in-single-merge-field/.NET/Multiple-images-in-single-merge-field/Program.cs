using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Linq;

namespace Multiple_images_in_single_merge_field
{
    class Program
    {
        static void Main(string[] args)
        {

            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Uses the mail merge events handler for image fields.
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ProductImage);
                    //Specifies the field names and field values.
                    string[] fieldNames = new string[] { "Logo" , "Name", "Company" };
                    string[] fieldValues = new string[] { "AdventureImages", "Nancy Davilo", "Syncfusion" };
                    //Performs the mail merge
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
		/// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge
            if (args.FieldName == "Logo")
            {
                //Gets the current merge field owner paragraph.
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                //Gets the current merge field index in the current paragraph.
                int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                //Gets the folder name from the field value.
                string ProductFolderName = args.FieldValue.ToString();
                // Define allowed image extensions
                string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff" };
                // Gets the image names from the Data folder
                string[] imageNames = Directory.GetFiles(Path.GetFullPath(@"Data/"))
                                               .Where(file => imageExtensions.Contains(Path.GetExtension(file).ToLower()))
                                               .ToArray();
                //Loops through the image names.
                foreach (string imageName in imageNames)
                {
                    //Gets the image from file system
                    FileStream imageStream = new FileStream(imageName, FileMode.Open, FileAccess.Read);
                    //Creates a new picture.
                    WPicture picture = new WPicture(paragraph.Document);
                    //Loads the image into picture.
                    picture.LoadImage(imageStream);
                    //Resizes the picture
                    picture.Height = 50;
                    picture.Width = 50;
                    //Inserts the picture at the current merge field index.
                    paragraph.ChildEntities.Insert(mergeFieldIndex, picture);
                    mergeFieldIndex++;
                }
                //Set field value as empty.
                args.Text = string.Empty;
            }
        }
    }
}
