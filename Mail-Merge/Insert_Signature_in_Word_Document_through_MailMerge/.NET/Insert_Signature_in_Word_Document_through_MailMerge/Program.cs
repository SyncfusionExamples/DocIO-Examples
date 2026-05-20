using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Insert_Signature_in_Word_Document_through_MailMerge
{
    class Program
    {

        static void Main(string[] args)
        {
            // Load the word document
            using (FileStream fileStream = new FileStream(Path.GetFullPath("../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    string[] fieldNames = { "Signature" };
                    string[] fieldValues = { "signature.png" };

                    document.MailMerge.MergeImageField += MailMerge_MergeSignature;
                    //Execute mail merge in the Word document
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    using (FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/Result.docx"), FileMode.Create, FileAccess.Write))
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
                
                WTextBox textbox = args.CurrentMergeField.OwnerParagraph.OwnerTextBody.Owner as WTextBox;
                // check whether the picture is inside the text box
                if (textbox != null)
                {
                    // Get the text box format
                    WTextBoxFormat textBoxFormat = textbox.TextBoxFormat;

                    if (textBoxFormat != null)
                    {
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
    }
}
