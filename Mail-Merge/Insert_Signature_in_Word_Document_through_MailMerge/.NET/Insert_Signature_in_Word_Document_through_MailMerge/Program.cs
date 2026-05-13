using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Insert_Signature_in_Word_Document_through_MailMerge
{
    class Program
    {

        static void Main(string[] args)
        {
            // Load the word document
            using (FileStream fileStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    string replacementText = "John’s Juice corner was established in the year of 2002 by John. Initially it was started in a small shop. Today Juice corner has over 300 branches over USA. The secret behind this success story is the recipes of John’s Mother Angelica. She has discovered about 500 secret recipes which are all used by John. ";

                    //Creates the bookmark navigator instance to access the bookmark
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark
                    bookmarksNavigator.MoveToBookmark("Bkmk");
                    //Replaces the bookmark content with text 
                    bookmarksNavigator.ReplaceBookmarkContent(replacementText, true);

                    string[] fieldNames = { "Signature" };
                    string[] fieldValues = { "signature.gif" };

                    document.MailMerge.MergeImageField += MailMerge_MergeSignature;
                    //Execute mail merge in the Word document
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        //Saves the stream as Word file
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Binds the image from file system and fit within text box during Mail merge process by using MergeImageFieldEventHandler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private static void MailMerge_MergeSignature(object sender, MergeImageFieldEventArgs args)
        {
            if (args.FieldName == "Signature")
            {
                string productFileName = args.FieldValue.ToString();
                byte[] imageBytes = File.ReadAllBytes(@"Data/" + productFileName);
                MemoryStream imageStream = new MemoryStream(imageBytes);
                args.ImageStream = imageStream;
                // Get the picture to be merged
                WPicture picture = args.Picture;              

                // Get the text box format
                WTextBoxFormat textBoxFormat = (args.CurrentMergeField.OwnerParagraph.OwnerTextBody.Owner as WTextBox).TextBoxFormat;

                // Resize width
                if (picture.Width != textBoxFormat.Width)
                {
                    float widthScale = textBoxFormat.Width / picture.Width * 100;
                    picture.WidthScale = widthScale;
                }

                // Resize height
                if (picture.Height != textBoxFormat.Height)
                {
                    float heightScale = textBoxFormat.Height / picture.Height * 100;
                    picture.HeightScale = heightScale;
                }
            }
        }
    }
}
